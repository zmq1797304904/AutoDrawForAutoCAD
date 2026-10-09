# -*- coding: utf-8 -*-
import cv2
import numpy as np
import re
import os
import sys
import ctypes
import shutil
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import time
import threading
import queue
import win32com.client
import pythoncom

# RapidOCR 引擎单例：模型只加载一次。首次运行会自动下载 ONNX 模型。
_OCR_ENGINE = None


class _OCREngineError(RuntimeError):
    """OCR 引擎层面的故障（如返回结构不兼容），区别于单张图片的识别失败。

    这类错误对每张图片都会重复出现，若被当作单图失败吞掉，用户只会看到
    "全部识别失败"而拿不到根因，因此 extract_coords 会把它原样抛出。
    """


# 绘图参数
CIRCLE_RADIUS = 0.5   # 圆半径 0.5m
CIRCLE_COLOR = 1      # AutoCAD 颜色索引：1=红(acRed)
# 圆心坐标映射：东坐标 -> X，北坐标 -> Y；平面绘图 Z 固定为 0
POINT_Z = 0.0

# 图层与高程文字标注参数
LAYER_CIRCLE = '长崖边巷道'             # 圆所在图层（不存在则自动创建）
LAYER_TEXT = '长崖边巷道高程文字标注'    # 高程文字所在图层（不存在则自动创建）
TEXT_HEIGHT = 0.8        # 文字高度
TEXT_WIDTH_FACTOR = 0.8  # 文字宽度因子
TEXT_COLOR = 3           # AutoCAD 颜色索引：3=绿(acGreen)
TEXT_STYLE = 'Standard'  # 文字样式
ZOOM_MARGIN = 5.0        # 绘制完成后视图缩放到绘制区域时的外扩边距（米）

# COM 瞬态错误重试参数：AutoCAD 忙时会拒绝 RPC 调用，需等待后重试
COM_RETRY_ATTEMPTS = 5   # 连接/打开图纸、单个坐标点绘制的重试次数
COM_ZOOM_ATTEMPTS = 3    # 视图缩放的重试次数
COM_RETRY_DELAY = 0.5    # 每次重试前的等待秒数

# 图层名的非法字符：AutoCAD 不接受这些符号，含则建图层必失败
_LAYER_ILLEGAL_CHARS = set('<>/\\":;?*|=,')

# 东/北坐标与高程的量级分界。地球最高点约 8848 m，而测量用的投影坐标即使采用
# 局部坐标系也基本在万米以上（实测东坐标 3.7e7、北坐标 3.5e6、高程约 350），
# 因此用 1e4 把高程与平距/高差之类的小数值同坐标区分开。
COORD_MIN_MAGNITUDE = 1e4

APP_USER_MODEL_ID = 'AutoDraw.AutoCAD.CoordinateTool'  # 任务栏归组与图标标识

# AutoCAD 颜色索引 (ACI) 到中文名映射，用于窗体下拉框显示；256=ByLayer（随图层）
_ACI_COLORS = [
    (1, '红'), (2, '黄'), (3, '绿'), (4, '青'),
    (5, '蓝'), (6, '紫'), (7, '白'), (8, '灰'),
    (256, 'ByLayer'),
]
_ACI_TO_NAME = {i: n for i, n in _ACI_COLORS}
_NAME_TO_ACI = {n: i for i, n in _ACI_COLORS}

# 主 root 引用（batch_process 创建后存到这里，便于 if __name__ 调用 mainloop
# 阻塞等用户关闭日志窗）
_ROOT = None
_LOG_REDIRECTOR = None  # 主线程日志重定向器，供后台线程安全调用 tkinter API
_PROGRESS_VAR = None    # 进度条 DoubleVar
_PROGRESS_LABEL = None  # 进度百分比 Label


def _resource_path(rel_name):
    """获取随程序分发的资源文件绝对路径。

    兼容两种运行方式：源码运行（脚本所在目录）与
    PyInstaller onefile 打包（sys._MEIPASS 临时解压目录）。
    """
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_name)


def _apply_window_icon(root):
    """为 tkinter 窗口设置程序图标，使任务栏/标题栏显示 AutoDraw.ico。

    需打包时通过 --add-data 将 AutoDraw.ico 嵌入；加载失败时静默跳过，
    不影响正常功能。
    """
    try:
        root.iconbitmap(_resource_path('AutoDraw.ico'))
    except Exception:
        pass


class _LogWindowRedirector:
    """将 print 输出重定向到日志窗体 Text 控件，同时保留原 stdout 输出。

    线程安全：处理逻辑在后台线程运行，print 产生的内容先放入线程安全队列，
    由主线程通过 root.after() 轮询取出并写入 Text 控件（tkinter 只能在主线程操作）。
    兼容打包后 --windowed 模式（无控制台）：原 stdout 可能为 None，写入失败时静默跳过。
    """

    def __init__(self, text_widget, root, original_stdout=None):
        self.text = text_widget
        self.root = root
        self.original = original_stdout
        self._queue = queue.Queue()
        self._ui_queue = queue.Queue()  # 需在主线程执行的任务（如 messagebox）
        # 主窗关闭时置位：唤醒所有阻塞在 run_on_main 的后台线程，避免永久挂起
        self._closed = threading.Event()
        # Text 控件最多保留的行数，超出后删除最前面的内容，防止长任务内存膨胀
        self._max_lines = 5000
        # 关键 emoji 前缀 -> Text tag 名（用于着色）
        self._tag_rules = (
            ('✅', 'ok'), ('❌', 'err'), ('⚠️', 'warn'),
            ('📝', 'info'), ('🔧', 'info'), ('🚀', 'info'),
            ('📌', 'info'), ('🎯', 'info'), ('📁', 'info'),
            ('📄', 'info'), ('⚙️', 'info'), ('💥', 'err'),
        )
        # 启动主线程轮询：每 50ms 批量取出队列中的日志写入 Text，并执行主线程任务
        self._poll()

    def close(self):
        """主窗关闭时调用：通知后台线程停止等待。"""
        self._closed.set()

    def run_on_main(self, func, *args, **kwargs):
        """在主线程执行 func 并阻塞等待返回值（线程安全）。

        用于后台线程调用 tkinter API（如 messagebox）。主窗关闭后立即抛出，
        避免后台线程永远阻塞在等待上。messagebox 需要等用户点击，故不设超时。
        """
        if self._closed.is_set():
            raise RuntimeError("主窗口已关闭，操作已取消")
        done = threading.Event()
        box = {}

        def _task():
            try:
                box['r'] = func(*args, **kwargs)
            except Exception as e:
                box['e'] = e
            finally:
                done.set()

        self._ui_queue.put(_task)
        while not done.wait(0.2):
            if self._closed.is_set():
                raise RuntimeError("主窗口已关闭，操作已取消")
        if 'e' in box:
            raise box['e']
        return box.get('r')

    def post_to_main(self, func):
        """把任务投递到主线程执行，不等待结果。

        进度条等无需返回值的场景使用，避免后台线程为每次更新做一次
        最长 50ms 的同步往返（大量点时累计可达数十秒）。
        """
        if not self._closed.is_set():
            self._ui_queue.put(func)

    def _poll(self):
        """主线程轮询：批量取出日志写入 Text，并执行主线程任务队列。"""
        try:
            while True:
                content = self._queue.get_nowait()
                self._append_to_text(content)
        except queue.Empty:
            pass
        # 执行需在主线程运行的任务（如 messagebox）
        try:
            while True:
                task = self._ui_queue.get_nowait()
                try:
                    task()
                except Exception:
                    # 单个任务失败不能中断轮询，否则所有后续日志与弹窗都会丢失
                    pass
        except queue.Empty:
            pass
        # 递归调度下一次轮询（仅在 root 仍存活时）
        try:
            self.root.after(50, self._poll)
        except Exception:
            pass

    def write(self, content):
        """后台线程调用：仅放入队列，不直接操作 UI。"""
        if content:
            self._queue.put(content)
        # 同步写原 stdout（IDE 控制台/打包后可能为 None）
        if self.original is not None:
            try:
                self.original.write(content)
            except Exception:
                pass

    def _append_to_text(self, content):
        """主线程调用：实际写入 Text 控件。"""
        if not content:
            return
        tag = ''
        for emoji, tag_name in self._tag_rules:
            if emoji in content:
                tag = tag_name
                break
        try:
            if tag:
                self.text.insert(tk.END, content, (tag,))
            else:
                self.text.insert(tk.END, content)
            # 长任务会把 Text 撑到几十万行，这里按行数上限裁剪最前面的内容
            if int(self.text.index('end-1c').split('.')[0]) > self._max_lines:
                self.text.delete('1.0', f'{self._max_lines // 5}.0')
            self.text.see(tk.END)
        except Exception:
            # 日志窗已被用户关闭：Text 控件已销毁
            pass

    def flush(self):
        if self.original is not None:
            try:
                self.original.flush()
            except Exception:
                pass


def _msg(func, title, msg, root):
    """线程安全的 messagebox 调用：后台线程通过此函数在主线程弹窗。"""
    if _LOG_REDIRECTOR is not None:
        return _LOG_REDIRECTOR.run_on_main(func, title, msg, parent=root)
    # 兜底（redirector 未初始化时直接调用，仅限主线程场景）
    return func(title, msg, parent=root)


def _set_progress(fraction):
    """线程安全地更新进度条。

    fraction: 0.0 ~ 1.0，进度条显示为百分比。
    主线程调用直接更新；后台线程通过 run_on_main 调度到主线程更新。
    """
    fraction = max(0.0, min(1.0, fraction))
    pct = int(round(fraction * 100))

    def _update():
        if _PROGRESS_VAR is not None:
            _PROGRESS_VAR.set(pct)
        if _PROGRESS_LABEL is not None:
            _PROGRESS_LABEL.config(text=f"{pct}%")

    # 主线程直接更新；后台线程投递给主线程（不等待，避免逐点同步往返累积延迟）
    if threading.current_thread() is threading.main_thread():
        _update()
    elif _LOG_REDIRECTOR is not None:
        _LOG_REDIRECTOR.post_to_main(_update)
    else:
        _update()


def _create_main_window():
    """创建主集成窗体：上半部分文件选择栏 + 下半部分日志输出区。

    窗体布局：
      顶栏 Frame（固定高度）：
        - 图片来源：单选按钮（图片文件夹 / Word 文档）+ 路径显示 Entry + 浏览按钮
        - 目标 DWG：路径显示 Entry + 浏览按钮
        - 开始处理按钮（路径选齐后启用）
      底栏 Frame（自动伸展）：
        - Text + Scrollbar（运行日志，print 输出实时显示）

    返回 (root, state)：root 是主窗口，state 是包含路径变量的字典。
    程序入口在 if __name__ 中调用 root.mainloop() 阻塞至用户关闭主窗。
    """
    root = tk.Tk()
    root.title("AutoDraw - 自动绘制坐标")
    root.minsize(700, 500)
    # 窗口水平+垂直居中显示（用固定尺寸计算，避免 winfo 在窗口未完全绘制时返回 1）
    _WIN_W, _WIN_H = 1000, 720
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    # Windows 任务栏占用底部空间，可用高度需减去约 50px（任务栏默认高度）
    _TASKBAR = 50
    avail_h = sh - _TASKBAR
    x = max(0, (sw - _WIN_W) // 2)
    y = max(0, (avail_h - _WIN_H) // 2)
    root.geometry(f"{_WIN_W}x{_WIN_H}+{x}+{y}")
    _apply_window_icon(root)

    # 处理中标志：防止重复启动 worker 线程，并让关窗时能提醒用户
    _busy = {'on': False}

    # 主窗关闭 = 程序退出
    def _on_close():
        import sys
        if _busy['on']:
            # 处理中关窗会直接切断进行中的 COM 调用，AutoCAD 可能留下半成品图形
            if not messagebox.askyesno(
                "正在处理中",
                "绘图任务尚未结束。强行退出可能让 AutoCAD 留下绘制到一半的图形。\n\n确定要退出吗？",
                parent=root,
            ):
                return
        sys.stdout = getattr(_create_main_window, '_orig_stdout', sys.__stdout__)
        # 通知后台线程停止等待主线程，否则 run_on_main 会永久挂起
        if _LOG_REDIRECTOR is not None:
            _LOG_REDIRECTOR.close()
        root.quit()
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", _on_close)

    # ---- 顶栏：文件选择 ----
    top = tk.LabelFrame(root, text="文件选择", font=("Microsoft YaHei", 10),
                        padx=10, pady=8)
    top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(8, 4))

    # 状态变量
    source_type = tk.StringVar(value='folder')    # 'folder' 或 'docx'
    folder_path = tk.StringVar(value='')          # 文件夹路径
    docx_paths_str = tk.StringVar(value='')       # docx 路径列表显示（多选以分号分隔）
    docx_paths = []                                # 实际路径元组
    dwg_path = tk.StringVar(value='')             # DWG 路径
    docx_paths_ref = docx_paths                    # 闭包共享引用

    # 行1：图片来源类型 + 路径 + 浏览
    row1 = tk.Frame(top)
    row1.pack(fill=tk.X, pady=(0, 6))

    tk.Label(row1, text="图片来源：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
    folder_radio = tk.Radiobutton(row1, text="图片文件夹", variable=source_type, value='folder',
                                  font=("Microsoft YaHei", 10),
                                  command=lambda: _update_browse_state())
    folder_radio.pack(side=tk.LEFT, padx=(4, 8))
    docx_radio = tk.Radiobutton(row1, text="Word 文档(.docx)", variable=source_type, value='docx',
                                font=("Microsoft YaHei", 10),
                                command=lambda: _update_browse_state())
    docx_radio.pack(side=tk.LEFT, padx=4)

    path_entry = tk.Entry(row1, font=("Microsoft YaHei", 10),
                          textvariable=folder_path if source_type.get() == 'folder' else docx_paths_str,
                          state='readonly')
    path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 4))

    browse_btn = tk.Button(row1, text="浏览...", width=8,
                           font=("Microsoft YaHei", 10),
                           command=lambda: _browse_source())
    browse_btn.pack(side=tk.LEFT)

    # 行2：目标 DWG
    row2 = tk.Frame(top)
    row2.pack(fill=tk.X, pady=(0, 8))

    tk.Label(row2, text="目标 DWG：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
    dwg_entry = tk.Entry(row2, font=("Microsoft YaHei", 10),
                         textvariable=dwg_path, state='readonly')
    dwg_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 4))
    dwg_browse_btn = tk.Button(row2, text="浏览...", width=8,
                               font=("Microsoft YaHei", 10),
                               command=lambda: _browse_dwg())
    dwg_browse_btn.pack(side=tk.LEFT)

    # 行3：开始处理按钮（居中）
    row3 = tk.Frame(top)
    row3.pack(fill=tk.X)
    start_btn = tk.Button(row3, text="▶  开始处理", font=("Microsoft YaHei", 11, "bold"),
                          height=1, width=20, command=lambda: _start_processing(),
                          state=tk.DISABLED)
    start_btn.pack()

    # ---- 中栏：绘图参数（可选，修改后覆盖默认值，留空或解析失败用默认）----
    param_frame = tk.LabelFrame(root, text="绘图参数（可选，留空用默认）",
                                font=("Microsoft YaHei", 10), padx=10, pady=6)
    param_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(0, 4))

    # 参数定义：(常量名, 标签, 默认值, 类型)；颜色用 'color' 类型走 Combobox 下拉框
    _param_defs = [
        ('CIRCLE_RADIUS', '圆半径', str(CIRCLE_RADIUS), float),
        ('CIRCLE_COLOR', '圆颜色', _ACI_TO_NAME[CIRCLE_COLOR], 'color'),
        ('LAYER_CIRCLE', '圆图层', LAYER_CIRCLE, str),
        ('TEXT_HEIGHT', '文字高度', str(TEXT_HEIGHT), float),
        ('TEXT_WIDTH_FACTOR', '文字宽度', str(TEXT_WIDTH_FACTOR), float),
        ('TEXT_COLOR', '文字颜色', _ACI_TO_NAME[TEXT_COLOR], 'color'),
        ('TEXT_STYLE', '文字样式', TEXT_STYLE, str),
        ('LAYER_TEXT', '文字图层', LAYER_TEXT, str),
        ('ZOOM_MARGIN', '缩放边距', str(ZOOM_MARGIN), float),
    ]
    param_vars = {}    # 常量名 -> StringVar
    param_types = {}   # 常量名 -> 类型
    param_widgets = {}  # 常量名 -> 输入控件（用于统一启停）
    # 数值参数的合理区间 (下界, 上界)，要求 下界 < 值 <= 上界
    _param_ranges = {
        'CIRCLE_RADIUS': (0.0, 1e6),
        'TEXT_HEIGHT': (0.0, 1e6),
        'TEXT_WIDTH_FACTOR': (0.0, 100.0),
        'ZOOM_MARGIN': (0.0, 1e6),
    }
    for i, (name, label, default, ptype) in enumerate(_param_defs):
        row = i // 3
        col = (i % 3) * 2
        tk.Label(param_frame, text=label, font=("Microsoft YaHei", 9),
                 width=8, anchor='e').grid(row=row, column=col, sticky='e', padx=(4, 2), pady=2)
        var = tk.StringVar(value=default)
        if ptype == 'color':
            # 颜色用 Combobox 下拉框显示中文名，readonly 防止用户输入任意值
            widget = ttk.Combobox(param_frame, textvariable=var,
                                  values=[n for _, n in _ACI_COLORS],
                                  width=12, state='readonly',
                                  font=("Microsoft YaHei", 9))
        else:
            widget = tk.Entry(param_frame, textvariable=var, width=14,
                              font=("Microsoft YaHei", 9))
        widget.grid(row=row, column=col + 1, sticky='we', padx=2, pady=2)
        param_vars[name] = var
        param_types[name] = ptype
        param_widgets[name] = widget
    # 列权重让 Entry 随窗口缩放伸展
    for c in range(6):
        param_frame.grid_columnconfigure(c, weight=1 if c % 2 else 0)

    # ---- 进度条（绘图参数与运行日志之间）----
    prog_frame = tk.Frame(root)
    prog_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(0, 4))
    tk.Label(prog_frame, text="进度：", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
    _progress_var = tk.DoubleVar(value=0.0)
    progress_bar = ttk.Progressbar(prog_frame, variable=_progress_var,
                                    maximum=100, mode='determinate')
    progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
    _progress_label = tk.Label(prog_frame, text="0%", font=("Microsoft YaHei", 9), width=6)
    _progress_label.pack(side=tk.LEFT, padx=(4, 0))
    global _PROGRESS_VAR, _PROGRESS_LABEL
    _PROGRESS_VAR = _progress_var
    _PROGRESS_LABEL = _progress_label

    # ---- 底栏：日志区（白色背景）----
    bot = tk.LabelFrame(root, text="运行日志", font=("Microsoft YaHei", 10),
                        padx=8, pady=6)
    bot.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))

    scrollbar = tk.Scrollbar(bot)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    log_text = tk.Text(bot, wrap=tk.CHAR, font=("Consolas", 10),
                       bg="#ffffff", fg="#1e1e1e", insertbackground="#1e1e1e",
                       yscrollcommand=scrollbar.set)
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=log_text.yview)

    # 颜色 tag（适配白色背景）
    log_text.tag_config('ok', foreground='#008000')     # 绿
    log_text.tag_config('err', foreground='#d00000')    # 红
    log_text.tag_config('warn', foreground='#b8860b')   # 暗黄
    log_text.tag_config('info', foreground='#0066cc')   # 蓝

    # ---- stdout 重定向到日志 Text ----
    import sys
    if not hasattr(_create_main_window, '_orig_stdout'):
        _create_main_window._orig_stdout = sys.stdout
    redirector = _LogWindowRedirector(log_text, root, _create_main_window._orig_stdout)
    sys.stdout = redirector
    global _LOG_REDIRECTOR
    _LOG_REDIRECTOR = redirector  # 存模块级引用，供后台线程安全调用 messagebox 等

    # ---- 回调：更新路径 Entry 关联的 StringVar ----
    def _update_browse_state():
        """切换来源类型时切换 path_entry 关联的 StringVar"""
        if source_type.get() == 'folder':
            path_entry.config(textvariable=folder_path)
        else:
            path_entry.config(textvariable=docx_paths_str)
        _check_start_ready()

    def _browse_source():
        """浏览：根据 source_type 弹 folder 或 docx 选择对话框"""
        if source_type.get() == 'folder':
            path = filedialog.askdirectory(title="请选择包含坐标照片的文件夹")
            if path:
                folder_path.set(path)
                docx_paths_str.set('')
                docx_paths_ref.clear()
        else:
            paths = filedialog.askopenfilenames(
                title="请选择包含坐标照片的 .docx 文档（可多选）",
                filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
            )
            if paths:
                docx_paths_ref.clear()
                docx_paths_ref.extend(paths)
                # 只显示文件名（多选以分号分隔），完整路径保留在 docx_paths_ref 中
                docx_paths_str.set("；".join(os.path.basename(p) for p in paths))
                folder_path.set('')
        _check_start_ready()

    def _browse_dwg():
        path = filedialog.askopenfilename(
            title="请选择要在上面绘制圆的 .dwg 文件",
            filetypes=[("DWG 文件", "*.dwg"), ("所有文件", "*.*")]
        )
        if path:
            dwg_path.set(path)
        _check_start_ready()

    def _check_start_ready():
        """路径齐全才启用开始按钮"""
        has_source = (source_type.get() == 'folder' and folder_path.get()) \
                     or (source_type.get() == 'docx' and docx_paths_ref)
        has_dwg = bool(dwg_path.get())
        start_btn.config(state=tk.NORMAL if (has_source and has_dwg) else tk.DISABLED)

    def _set_controls_enabled(enabled):
        """统一启停输入控件。

        readonly 的 Entry/Combobox 恢复时必须还原为 readonly：原先一律设成
        NORMAL 会把路径框和颜色下拉框解锁成可自由编辑。
        start_btn 不在此处处理，它的可用性取决于路径是否选齐。
        """
        state = tk.NORMAL if enabled else tk.DISABLED
        browse_btn.config(state=state)
        dwg_browse_btn.config(state=state)
        folder_radio.config(state=state)
        docx_radio.config(state=state)
        # 注意：tkinter 没有 tk.READONLY 常量，只读态就是字符串 'readonly'
        path_entry.config(state='readonly' if enabled else tk.DISABLED)
        dwg_entry.config(state='readonly' if enabled else tk.DISABLED)
        for name, widget in param_widgets.items():
            try:
                widget.config(state=('readonly' if param_types[name] == 'color'
                                     else tk.NORMAL) if enabled else tk.DISABLED)
            except tk.TclError:
                pass

    def _restore_ui():
        """恢复所有控件为可用状态（在主线程执行）。"""
        _busy['on'] = False
        _set_controls_enabled(True)
        start_btn.config(text="▶  开始处理")
        _check_start_ready()
        print("\n提示：可重新选择路径并继续处理下一批，或关闭窗口退出程序。")

    def _start_processing():
        """点击开始：禁用按钮，后台线程运行 batch_process，结束后恢复控件。

        处理在后台线程执行，主线程保持响应（窗口可拖动、日志实时刷新）。
        """
        # 重入守卫：控件禁用状态可能被用户中途操作绕开，这里做最终防线，
        # 防止两个 worker 线程同时抢同一张 DWG、同一个日志区和进度条
        if _busy['on']:
            print("⚠️ 已有处理任务在运行，请等待本次处理结束。")
            return
        folder = folder_path.get() if source_type.get() == 'folder' else None
        docx = tuple(docx_paths_ref) if source_type.get() == 'docx' else ()
        dwg = dwg_path.get()
        # 清空日志区，避免上一次的输出残留（保持 NORMAL 状态：DISABLED 会阻止程序 insert）
        log_text.delete('1.0', tk.END)
        # 重置进度条为 0%
        _set_progress(0.0)
        # 应用绘图参数：读取 Entry 值，按类型转换更新模块级常量
        # 解析失败或留空时保留原默认值，仅打印警告
        _apply_params()
        # 禁用所有控件，防止处理过程中误操作
        _set_controls_enabled(False)
        start_btn.config(state=tk.DISABLED, text="处理中...")
        root.update_idletasks()

        def _worker():
            """后台线程：执行 batch_process，结束后在主线程恢复 UI。"""
            # 后台线程必须初始化 COM 库（STA 模式），否则 win32com 创建的
            # AutoCAD COM 代理会失效，出现 "<unknown>.Layers" 等错误
            pythoncom.CoInitialize()
            try:
                batch_process(folder, docx, dwg, root)
            except Exception as e:
                # 用 print 而非 traceback.print_exc()：后者写 sys.stderr，
                # 而 --windowed 打包后 sys.stderr 可能为 None，会在异常处理里再抛一次
                import traceback
                print(f"\n💥 程序异常终止: {e}")
                print(traceback.format_exc())
            finally:
                # 恢复 UI 必须在主线程（tkinter 非线程安全）
                try:
                    if _LOG_REDIRECTOR is not None:
                        _LOG_REDIRECTOR.run_on_main(_restore_ui)
                    else:
                        _restore_ui()
                except RuntimeError:
                    pass  # 主窗已关闭，无需也无法恢复 UI
                except Exception as e:
                    # 恢复失败会让控件永久卡在禁用状态，必须让用户看到原因而不是静默吞掉
                    import traceback
                    print(f"⚠️ 恢复界面状态失败: {e}")
                    print(traceback.format_exc())
                # 与 CoInitialize 配对释放，避免多轮处理累积泄漏 COM 单元
                pythoncom.CoUninitialize()

        _busy['on'] = True
        threading.Thread(target=_worker, daemon=True).start()

    def _apply_params():
        """读取参数 Entry/Combobox，按声明类型转换、校验后更新模块级常量。
        留空或校验不通过时保留原默认值，仅打印警告。"""
        module_globals = globals()
        for name, var in param_vars.items():
            raw = var.get().strip()
            if not raw:
                continue
            ptype = param_types[name]
            old = module_globals.get(name)
            try:
                if ptype == 'color':
                    # 颜色：从中文名查 ACI 索引
                    new_val = _NAME_TO_ACI.get(raw)
                    if new_val is None:
                        raise ValueError("未知颜色名")
                elif ptype is str:
                    new_val = raw
                    if name.startswith('LAYER_'):
                        bad = _LAYER_ILLEGAL_CHARS & set(new_val)
                        if bad:
                            raise ValueError(f"图层名不能包含 {' '.join(sorted(bad))}")
                else:
                    new_val = ptype(raw)
                    # 半径为 0 会画出不可见圆；文字高度为 0 会让 AddText 直接抛异常
                    rng = _param_ranges.get(name)
                    if rng is not None and not (rng[0] < new_val <= rng[1]):
                        raise ValueError(f"需满足 {rng[0]} < 值 <= {rng[1]}")
                if new_val != old:
                    module_globals[name] = new_val
                    # 颜色项额外显示 ACI 与颜色名，便于核对
                    display = f"{new_val}({_ACI_TO_NAME.get(new_val, '?')})" if ptype == 'color' else new_val
                    print(f"⚙️ 参数已更新: {name} = {display}（默认 {old}）")
            except (ValueError, TypeError) as e:
                print(f"⚠️ 参数 {name}='{raw}' 不合法（{e}），保留默认 {old}")

    _update_browse_state()
    root.update()
    return root, {'folder_path': folder_path, 'docx_paths': docx_paths_ref,
                 'dwg_path': dwg_path, 'source_type': source_type}


def _trunc3(x):
    """截断到小数点后 3 位（不四舍五入），返回字符串。

    用字符串切片而非 round()，避免 343.4956 被四舍五入成 343.496。
    取 10 位小数再切片，规避浮点数精度导致的截断错位。
    """
    s = f"{float(x):.10f}"
    dot = s.index('.')
    s = s[:dot + 4]
    # 负高程加上负偏移后可能截断成 "-0.000"，去掉无意义的负号
    if s.startswith('-') and float(s) == 0:
        s = s[1:]
    return s


def _get_ocr_engine():
    """延迟初始化 RapidOCR，避免未选文件就加载模型。"""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        from rapidocr import RapidOCR
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _load_bgr(image_path):
    """用 numpy 读文件再解码，兼容中文路径。"""
    return cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)


def _enhance_for_blur(img):
    """备用通道：放大 + 锐化（仍保持三通道，适配检测模型）。"""
    h, w = img.shape[:2]
    up = cv2.resize(img, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(up, cv2.COLOR_BGR2GRAY)
    sharpened = cv2.addWeighted(gray, 1.6, cv2.GaussianBlur(gray, (0, 0), 2), -0.6, 0)
    return cv2.cvtColor(sharpened, cv2.COLOR_GRAY2BGR)


def _reading_order_texts(result):
    """按从上到下、从左到右排列识别结果，避免检测顺序打乱东/北/高程。"""
    txts = getattr(result, "txts", None)
    boxes = getattr(result, "boxes", None)
    if not txts:
        return []
    if boxes is None:
        return list(txts)

    keyed = []
    for box, txt in zip(boxes, txts):
        arr = np.asarray(box, dtype=np.float64)
        if arr.ndim == 2:
            y = float(arr[:, 1].min())
            x = float(arr[:, 0].min())
        else:
            y = float(arr.reshape(-1)[1])
            x = float(arr.reshape(-1)[0])
        keyed.append((y, x, txt))
    keyed.sort(key=lambda item: (item[0], item[1]))
    return [item[2] for item in keyed]


def _normalize_ocr_text(text):
    """全角数字/小数点转半角，便于后续正则提取。"""
    trans = str.maketrans("０１２３４５６７８９．－", "0123456789.-")
    return text.translate(trans)


def _ocr_to_text(img):
    result = _get_ocr_engine()(img)
    # RapidOCR 3.x 返回 RapidOCROutput（dataclass，字段 boxes/txts）。若上游改了返回
    # 结构，getattr 会静默取到 None，表现为"每张图都识别失败"却毫无报错，极难定位，
    # 因此这里显式检查属性是否存在——缺属性说明是引擎层面故障而非单图识别失败。
    if not (hasattr(result, 'txts') and hasattr(result, 'boxes')):
        raise _OCREngineError(
            f"OCR 返回结构异常（{type(result).__name__}，缺少 txts/boxes），"
            "请检查 requirements.txt 中 rapidocr 的版本约束"
        )
    return _normalize_ocr_text("\n".join(_reading_order_texts(result)))


def _decimal_numbers(text):
    """按出现顺序返回文本中所有带小数点的数字（字符串形式）。

    不做量级过滤：高程可能只有几十（只要大于 0），过滤会把它丢掉，
    导致后续"取最后三个数"整体错位。量级判断交给 _split_coords。
    """
    return re.findall(r"-?\d+\.\d+", text)


def _split_coords(numbers):
    """从数字列表尾部解析出 (东坐标, 北坐标, 高程)，无法判定时返回 None。

    仪器屏幕自上而下的顺序是 平距 -> 高差 -> 东坐标 -> 北坐标 -> 高程，
    因此取最后三个数：末位是高程，其前两位依次是北坐标、东坐标。

    只按位置取数是不够的：屏幕上若多出一个小数（如换页残留、OCR 误识的
    杂值），末位就会被误当成高程，整体错位后画出错误的点。因此额外校验
    高程与坐标的量级；不通过时返回 None，让调用方按"识别失败"处理——
    画错点比不画更危险。
    """
    if len(numbers) < 3:
        return None
    try:
        easting = float(numbers[-3])
        northing = float(numbers[-2])
        elevation = float(numbers[-1])
    except (TypeError, ValueError):
        return None

    if not 0 < elevation < COORD_MIN_MAGNITUDE:
        return None
    if abs(easting) < COORD_MIN_MAGNITUDE or abs(northing) < COORD_MIN_MAGNITUDE:
        return None
    return easting, northing, elevation


# 备注解析：位置后缀关键词
_POSITION_KEYWORDS = ('中顶', '右顶', '左顶')


def _parse_elevation_note(text):
    """解析 docx 段落文字备注，返回 (offset, suffix) 或 None。

    支持三种形式（同段可多次出现并累积）：
      - 纯数值："-0.1"、"+1.3"、"-1"   → 高程增减
      - 纯位置："中顶"、"右顶"、"左顶" → 文字标注追加后缀
      - 组合：  "+0.6中顶"、"-0.2右顶" → 数值修正 + 后缀追加

    过滤策略：|offset| > 5 视为非备注（如段落编号、日期等）。
    后缀去重：同段重复"中顶中顶"只保留一个。
    """
    if not text:
        return None

    offset = 0.0
    suffix_parts = []

    # 1. 组合形式（数值+位置）：优先匹配，避免被纯数值规则吞掉后缀
    combo_pat = re.compile(r'([+-]?\d+\.?\d*)\s*(中顶|右顶|左顶)')
    consumed_spans = []  # 已被组合规则消费的字符区间，避免纯数值规则二次匹配
    for m in combo_pat.finditer(text):
        try:
            v = float(m.group(1))
        except ValueError:
            continue
        if abs(v) > 5:
            continue
        offset += v
        if m.group(2) not in suffix_parts:
            suffix_parts.append(m.group(2))
        consumed_spans.append((m.start(), m.end()))

    # 2. 纯位置关键词（未被组合规则覆盖的）
    pos_pat = re.compile(r'(中顶|右顶|左顶)')
    consumed_pos_spans = []
    for m in pos_pat.finditer(text):
        # 落在已消费区间内则跳过
        if any(s <= m.start() and m.end() <= e for s, e in consumed_spans):
            continue
        if m.group(1) not in suffix_parts:
            suffix_parts.append(m.group(1))
        consumed_pos_spans.append((m.start(), m.end()))

    # 3. 纯数值（未被组合规则覆盖的）：|v|<=5 才接受
    num_pat = re.compile(r'([+-]?\d+\.?\d*)')
    for m in num_pat.finditer(text):
        if any(s <= m.start() and m.end() <= e for s, e in consumed_spans):
            continue
        try:
            v = float(m.group(1))
        except ValueError:
            continue
        if abs(v) > 5:
            continue
        # 过滤明显非备注：整数部分 > 2 位（如段落编号"2024"、日期"20240301"）
        int_part = m.group(1).lstrip('+-').split('.')[0]
        if len(int_part) > 2:
            continue
        offset += v

    if offset == 0.0 and not suffix_parts:
        return None
    return (offset, ''.join(suffix_parts))


def extract_coords(image_path):
    """从单张图片中提取坐标。

    仪器屏幕的数值排列顺序为：东坐标 -> 北坐标 -> 高程，
    因此取最后三个数依次作为 东坐标、北坐标、高程（详见 _split_coords）。
    """
    try:
        img = _load_bgr(image_path)
        if img is None:
            return {"文件名": os.path.basename(image_path), "东坐标": "识别失败",
                    "北坐标": "图片无法读取或已损坏", "高程": "识别失败"}

        # 先用原图识别；结果不可信（数量不足或校验不通过）时再用放大锐化图重试
        coords = None
        for candidate in (img, _enhance_for_blur(img)):
            coords = _split_coords(_decimal_numbers(_ocr_to_text(candidate)))
            if coords is not None:
                break

        if coords is not None:
            easting, northing, elevation = coords
            return {
                "文件名": os.path.basename(image_path),
                "东坐标": _trunc3(easting),
                "北坐标": _trunc3(northing),
                "高程": _trunc3(elevation),
            }
        return {"文件名": os.path.basename(image_path), "东坐标": "识别失败", "北坐标": "识别失败",
                "高程": "识别失败"}

    except _OCREngineError:
        raise  # 引擎级故障：中断整轮处理，避免每张图重复记录同一条根因
    except Exception as e:
        return {"文件名": os.path.basename(image_path), "东坐标": f"报错: {e}", "北坐标": "",
                "高程": ""}


def _ensure_layer(doc, layer_name):
    """获取指定图层，不存在则自动创建。

    Item 可能因 AutoCAD 忙被瞬态拒绝，紧接着的 Add 又会以"图层已存在"失败，
    因此 Add 失败后再取一次 Item 兜底；两次都失败才真正抛出。
    """
    try:
        return doc.Layers.Item(layer_name)
    except Exception:
        pass
    try:
        return doc.Layers.Add(layer_name)
    except Exception:
        return doc.Layers.Item(layer_name)


def _is_com_call_rejected(exc):
    """判断异常是否为可重试的瞬态 COM 错误。

    AutoCAD 处于忙状态（重绘、打开图纸、用户操作或对话框处理）时会拒绝 COM 调用，
    短暂等待后重试通常即可成功。识别三种表现形式：
      - RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER（有明确 hresult）
      - win32com 偶发地把属性解析为方法："Property 'X' can not be set"
      - 晚绑定拿不到类型信息时的成员名报错："<unknown>.Layers"
    """
    hr = None
    if hasattr(exc, 'hresult'):
        hr = exc.hresult
    elif getattr(exc, 'args', None):
        try:
            hr = int(exc.args[0])
        except (TypeError, ValueError):
            pass
    # 可重试的瞬态 COM 错误：
    #   RPC_E_CALL_REJECTED      (0x80010001) -2147418111  "被呼叫方拒绝接收呼叫"
    #   RPC_E_SERVERCALL_RETRYLATER (0x8001010A) -2147417846  "消息筛选器显示应用程序正在使用中"
    if hr in (-2147418111, -2147417846):
        return True
    # 偶发 win32com 属性解析错误（"Property 'X' can not be set"）
    msg = str(exc)
    if "can not be set" in msg:
        return True
    # 晚绑定时 pywin32 拿不到类型信息，成员名解析失败的报错会被格式化成
    # "<unknown>.Layers" 或 "AddText.Update" 这类 "对象名.成员名" 文本
    # （本项目实测出现在 AutoCAD 忙、拒绝 COM 调用时）。
    # 这种形式取不到有意义的 hresult，只能按文本判断。
    # 若真是"成员不存在"，重试几次后仍会以同样的错误失败，代价只是几秒等待。
    if msg.startswith("<unknown>."):
        return True
    # 匹配 "ClassName.MemberName" 形式（如 "AddText.Update"、"AddCircle.color"），
    # 这些都是 win32com 晚绑定拿不到类型信息时的典型报错格式
    import re
    if re.match(r'^[A-Z][A-Za-z0-9_]*\.[A-Za-z0-9_]+$', msg):
        return True
    return False


def _com_wait(seconds):
    """在当前 STA 线程内等待指定秒数，期间持续泵送消息队列。

    AutoCAD 忙时对被拒的调用需要调用方重新发起才可能成功，而 STA 线程长时间
    不泵消息也会让排队中的调用得不到处理。这里按 COM 的常规做法边等待边
    PumpWaitingMessages，避免单纯 sleep 把消息处理一起阻塞掉。
    """
    deadline = time.time() + seconds
    while True:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            # 非 STA 线程或消息过滤器异常：退化为纯等待，不影响功能
            pass
        remaining = deadline - time.time()
        if remaining <= 0:
            return
        time.sleep(min(0.05, remaining))


def _wait_until_ready(doc, timeout=10.0):
    """等待 AutoCAD 完成图纸激活/重生成；未就绪则抛出最后一次拒绝异常。

    打开或激活图纸后 AutoCAD 仍会忙一段时间（大图纸重生成更久，人工操作或
    弹出的对话框也会挡在前面），期间所有 COM 调用都被拒绝。原先固定睡 1 秒
    并不可靠——大图纸上绘制阶段的第一个调用就会撞上忙窗口。
    这里用一个轻量属性读取探测就绪，就绪即返回；非瞬态错误立即抛出；
    超时则把最后一次拒绝抛给调用方，由其重试逻辑处理。
    """
    deadline = time.time() + timeout
    last_err = None
    while True:
        try:
            doc.Name  # 轻量读取：能成功即说明 AutoCAD 已可响应
            return
        except Exception as e:
            if not _is_com_call_rejected(e):
                raise
            last_err = e
        if time.time() >= deadline:
            print(f"   ⚠️ AutoCAD 等待 {timeout:.0f} 秒仍未就绪，重试中...")
            raise last_err
        _com_wait(COM_RETRY_DELAY)


def _preflight_draw_env(doc):
    """绘制前预检：确认图层可创建、文字样式存在。

    参数不合法时若等到逐点绘制才暴露，每个点都要重试满 COM_RETRY_ATTEMPTS 次，
    成百个点会白白空转数分钟，因此提前一次性报出。
    """
    _ensure_layer(doc, LAYER_CIRCLE)
    _ensure_layer(doc, LAYER_TEXT)
    doc.TextStyles.Item(TEXT_STYLE)


def _draw_on_doc(doc, points):
    """在给定文档上绘制圆与高程文字标注。

    points: [(东坐标X, 北坐标Y, 高程, 后缀, 文件名), ...]
    返回 (成功数量, 失败列表[(文件名, 失败原因)])。
    单个点绘制失败不影响其他点；若圆或文字只画出一半，会把已创建的图元
    一并删除，保持整组回滚，避免图上留下没有标注的孤圆或孤立文字。
    """
    # 绘制前的第一次 COM 接触最容易撞上 AutoCAD 忙窗口（被 RPC_E_CALL_REJECTED
    # 或以 "<unknown>.Layers" 形式拒绝）。原先这里没有任何重试，一次瞬态拒绝
    # 就让整批坐标点全部失败，因此按与其他 COM 阶段一致的方式重试。
    preflight_err = None
    for attempt in range(COM_RETRY_ATTEMPTS):
        try:
            _preflight_draw_env(doc)
            preflight_err = None
            break
        except Exception as e:
            preflight_err = e
            if _is_com_call_rejected(e) and attempt < COM_RETRY_ATTEMPTS - 1:
                _com_wait(COM_RETRY_DELAY)
                continue
            break
    if preflight_err is not None:
        reason = str(preflight_err).replace('\n', ' ')[:200]
        raise RuntimeError(f"绘图前无法访问图纸对象（AutoCAD 可能未响应）: {reason}")

    model_space = doc.ModelSpace
    drawn = 0
    failed = []
    total_pts = len(points)
    for pi, (easting, northing, elevation, suffix, source_file) in enumerate(points):
        # AutoCAD 处于忙状态时会拒绝 COM 调用 (RPC_E_CALL_REJECTED)，
        # 单个点的绘制流程作为 retry 单元：失败时清理半成品后短暂等待重试。
        # 非 RPC 拒绝错误直接记录失败不重试。
        last_err = None
        for attempt in range(COM_RETRY_ATTEMPTS):
            circle = None
            text_obj = None
            try:
                # ---- 红色圆（图层：长崖边巷道）----
                center = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8,
                    (float(easting), float(northing), POINT_Z)
                )
                circle = model_space.AddCircle(center, CIRCLE_RADIUS)
                circle.Layer = LAYER_CIRCLE
                circle.color = CIRCLE_COLOR   # 红色
                circle.Update()

                # ---- 高程文字（图层：长崖边巷道高程文字标注，绿色，与圆重叠居中）----
                # 高程截断 3 位后拼接后缀（如 "1157.225中顶"）
                text_content = _trunc3(elevation) + suffix
                text_obj = model_space.AddText(text_content, center, TEXT_HEIGHT)
                text_obj.Layer = LAYER_TEXT
                text_obj.color = TEXT_COLOR      # 绿色
                text_obj.StyleName = TEXT_STYLE
                text_obj.ScaleFactor = TEXT_WIDTH_FACTOR
                text_obj.HorizontalAlignment = 4  # acHorizontalAlignmentMiddle：以对齐点为文字中心
                text_obj.TextAlignmentPoint = center
                text_obj.Update()

                drawn += 1
                print(f"   ✅ 已绘制圆+高程标注: X(东)={easting}  Y(北)={northing}  高程={text_content}  <- {source_file}")
                last_err = None
                break  # 成功，跳出 retry 循环
            except Exception as e:
                # 只成功了一半：把已创建的图元全部删除再重试，
                # 否则失败点会在图上留下孤圆或孤立文字（文字可能先于圆之后创建成功）
                for obj in (text_obj, circle):
                    if obj is not None:
                        try:
                            obj.Delete()
                        except Exception:
                            pass
                last_err = e
                if _is_com_call_rejected(e) and attempt < COM_RETRY_ATTEMPTS - 1:
                    # RPC 调用被拒绝：泵消息等待 AutoCAD 处理完手头工作再 retry
                    _com_wait(COM_RETRY_DELAY)
                    continue
                # 非重试型错误（或已用尽重试次数）：直接记录失败
                break

        if last_err is not None:
            reason = str(last_err).replace('\n', ' ')[:200]
            failed.append((source_file, f"绘制失败: {reason}"))
            print(f"   ❌ 绘制失败: {source_file} -> {reason}")
        # 绘图阶段占总进度的后 50%（50% ~ 100%）
        _set_progress(0.5 + (pi + 1) / total_pts * 0.5)

    return drawn, failed


def draw_circles_on_dwg(dwg_path, points):
    """打开指定 DWG 文件，在每个坐标点绘制红色圆并标注高程文字。

    points: [(东坐标X, 北坐标Y, 高程, 文件名), ...]
    坐标系：东坐标 -> CAD 的 X 轴，北坐标 -> CAD 的 Y 轴
    """
    abs_dwg = os.path.abspath(dwg_path)

    # 连接/打开图纸阶段也可能因 AutoCAD 忙被拒（RPC_E_CALL_REJECTED），
    # 整体包在 retry 循环内；非瞬态错误立即抛出
    last_err = None
    for attempt in range(COM_RETRY_ATTEMPTS):
        try:
            # 连接（或启动）AutoCAD
            acad = win32com.client.Dispatch("AutoCAD.Application")
            acad.Visible = True

            # 若该图纸已在 AutoCAD 中打开则直接激活，否则重新打开
            doc = None
            for open_doc in acad.Documents:
                try:
                    if os.path.abspath(open_doc.FullName).lower() == abs_dwg.lower():
                        doc = open_doc
                        break
                except Exception:
                    continue
            if doc is None:
                # win32com 动态分发下 Open 的返回值不可靠（会拿到方法包装对象），
                # 显式标记为方法后调用，再从 ActiveDocument / Documents 集合取回文档对象
                acad.Documents._FlagAsMethod("Open")
                acad.Documents.Open(abs_dwg)
                for _ in range(60):
                    try:
                        active = acad.ActiveDocument
                        if os.path.abspath(active.FullName).lower() == abs_dwg.lower():
                            doc = active
                            break
                    except Exception:
                        pass
                    _com_wait(1.0)
                if doc is None:
                    for open_doc in acad.Documents:
                        try:
                            if os.path.abspath(open_doc.FullName).lower() == abs_dwg.lower():
                                doc = open_doc
                                break
                        except Exception:
                            continue
            # doc 仍为 None 时若继续调用 doc.Activate()，抛出的只是
            # "'NoneType' object has no attribute 'Activate'"，用户完全看不出
            # 是图纸打开超时，这里换成带业务含义的报错
            if doc is None:
                raise RuntimeError(
                    f"打开图纸超时（60 秒内未完成）：{os.path.basename(abs_dwg)}。"
                    "请确认文件未被占用、路径可访问，且 AutoCAD 未弹出对话框等待响应。"
                )
            doc.Activate()
            # 等待 AutoCAD 完成激活/重生成再进入绘制。原先固定睡 1 秒不够：
            # 大图纸激活后仍在忙，绘制阶段的第一个 COM 调用就会撞上忙窗口被拒绝
            _wait_until_ready(doc)
            break  # 连接+打开成功，跳出 retry 循环
        except Exception as e:
            last_err = e
            if _is_com_call_rejected(e):
                # AutoCAD 忙，等待后重试
                _com_wait(COM_RETRY_DELAY)
                continue
            raise  # 非瞬态错误直接抛出
    else:
        # 全部重试都被拒绝，抛出最后一次异常
        raise last_err

    drawn, failed = _draw_on_doc(doc, points)

    # 视图缩放到本次成功绘制的区域（外扩 ZOOM_MARGIN 米），而不是整张图，
    # 避免 ZoomExtents 把大坐标图纸缩得过小、难以确认刚绘制的点
    # 注意 points 元素为 (东, 北, 高程, 后缀, 文件名)，文件名是下标 4
    failed_names = {name for name, _ in failed}
    success_points = [pt for pt in points if pt[4] not in failed_names]
    if success_points:
        xs = [pt[0] for pt in success_points]
        ys = [pt[1] for pt in success_points]
        margin = ZOOM_MARGIN
        lower = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (min(xs) - margin, min(ys) - margin, POINT_Z)
        )
        upper = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (max(xs) + margin, max(ys) + margin, POINT_Z)
        )
        # ZoomWindow 同样可能被 AutoCAD 拒绝（RPC_E_CALL_REJECTED），加 retry；
        # 失败时仅打印警告不抛出异常——前面已成功绘制的圆和文字标注不受影响
        for attempt in range(COM_ZOOM_ATTEMPTS):
            try:
                acad.ZoomWindow(lower, upper)
                break
            except Exception as e:
                if _is_com_call_rejected(e) and attempt < COM_ZOOM_ATTEMPTS - 1:
                    _com_wait(COM_RETRY_DELAY)
                    continue
                # 重试用尽或非重试错误：缩放失败不影响已绘制内容
                print(f"   ⚠️ 视图缩放失败（已绘制内容不受影响）: {e}")
                break
    return drawn, failed


def archive_failed_images(folder_path, failures):
    """把失败图片复制到照片目录下带时间戳的新文件夹，并写入失败原因清单。

    failures: [(图片完整路径, 文件名, 失败原因), ...]
    返回归档文件夹路径；无失败项或归档不可用时返回 None。

    归档只是"方便人工复查"的附带动作，不应中断整轮处理：目录只读、网络盘掉线、
    磁盘写满等都会让这里抛异常，若冒泡出去会导致用户连汇总弹窗都看不到，
    因此内部异常一律降级为警告。
    """
    if not failures or not folder_path:
        return None

    archive_dir = os.path.join(
        folder_path, f"识别失败图片_{datetime.now():%Y%m%d_%H%M%S}"
    )
    try:
        os.makedirs(archive_dir, exist_ok=True)
    except Exception as e:
        print(f"   ⚠️ 无法创建失败归档目录（{archive_dir}）: {e}")
        return None

    used_names = set()
    for full_path, filename, reason in failures:
        # 防止重名文件互相覆盖
        dest_name = filename
        if dest_name in used_names:
            stem, ext = os.path.splitext(filename)
            dest_name = f"{stem}_{len(used_names)}{ext}"
        used_names.add(dest_name)
        try:
            shutil.copy2(full_path, os.path.join(archive_dir, dest_name))
        except Exception as e:
            print(f"   ⚠️ 复制失败图片出错: {filename} -> {e}")

    # 写出失败原因清单，便于后续人工分析
    try:
        log_path = os.path.join(archive_dir, "失败原因.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"失败数量: {len(failures)}\n")
            f.write("=" * 60 + "\n")
            for _, filename, reason in failures:
                f.write(f"{filename}\t{reason}\n")
    except Exception as e:
        print(f"   ⚠️ 写入失败原因清单出错: {e}")

    return archive_dir


def extract_images_from_docx(docx_path, output_dir):
    """从 docx 文档中按出现顺序提取所有嵌入式图片到 output_dir，并解析每张图片上方紧挨的文字备注。

    返回 [(保存后的完整路径, 显示用文件名, 备注字符串), ...]。
    备注字符串为该图上方连续文字段落的合并文本（可能为 ""）。
    文件名格式：{文档名(去扩展名)}_{序号:03d}.{扩展名}，便于追溯来源。

    解析策略：
    1. 优先用 python-docx 遍历段落（保留图片与文字的相对顺序，能解析备注）；
    2. 若文档结构不标准（如 WPS 生成的 docx 缺少 docProps/core.xml 导致
       python-docx 无法打开），则回退到直接用 zipfile 解压 word/media/，
       按文件名序号排序提取；此时无法解析备注，note 全部为 ""。
    """
    from docx import Document
    from docx.oxml.ns import qn
    import hashlib

    doc_base = os.path.splitext(os.path.basename(docx_path))[0]
    # 不同目录下可能有同名 docx（如两份"图片备忘.doc"），仅用文档名做前缀会让
    # 后一个文档的图片覆盖前一个，这里附加路径哈希保证跨文档唯一
    doc_tag = hashlib.md5(os.path.abspath(docx_path).encode('utf-8')).hexdigest()[:6]

    # 内容类型到扩展名的映射：只收 OCR 能解码的位图。
    # emf/wmf 是 Word 里常见的矢量图，cv2 无法解码，提取出来只会变成
    # 一条"图片无法读取"的失败记录，因此不收。
    ext_map = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/jpg': '.jpg',
        'image/gif': '.gif',
        'image/bmp': '.bmp',
        'image/tiff': '.tiff',
    }
    # zip 内扩展名到 content_type 的反向映射（回退模式用）
    zip_ext_map = {
        '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
        '.gif': 'image/gif', '.bmp': 'image/bmp', '.tiff': 'image/tiff',
    }

    def _save_image(blob, content_type, idx, note):
        """保存一张图片，返回 (路径, 文件名, 备注)；不支持的格式返回 None。"""
        ext = ext_map.get(content_type)
        if ext is None:
            # 不谎报成 .png：下游会按扩展名判断格式，错的扩展名会让问题更隐蔽
            print(f"   ⚠️ 跳过不支持的图片类型 {content_type}"
                  f"（{os.path.basename(docx_path)} 第 {idx} 张）")
            return None
        filename = f"{doc_base}_{doc_tag}_{idx:03d}{ext}"
        save_path = os.path.join(output_dir, filename)
        try:
            with open(save_path, 'wb') as f:
                f.write(blob)
        except Exception as e:
            print(f"   ⚠️ 保存图片失败 {filename}: {e}")
            return None
        return (save_path, filename, note)

    # ---- 策略1：python-docx 遍历段落（保留顺序、解析备注）----
    try:
        doc = Document(docx_path)
        extracted = []
        idx = 0
        # 待挂载备注缓冲区：遇到图片时把缓冲区内容作为该图片的 note，然后清空
        pending_notes = []

        for para in doc.paragraphs:
            # 提取段落内所有图片（inline + anchor 都用 <a:blip r:embed="rIdX">）
            blips = para._p.findall('.//' + qn('a:blip'))
            para_text = para.text.strip()

            if blips:
                # 段落含图片：把已累积的备注挂到本段第一张图上，
                # 同段多张图按顺序继承（同段多图共用同一组备注）
                note_for_first = '\n'.join(pending_notes)
                pending_notes.clear()
                for bi, blip in enumerate(blips):
                    rId = blip.get(qn('r:embed'))
                    if not rId:
                        continue
                    try:
                        part = doc.part.related_parts[rId]
                    except Exception:
                        continue
                    idx += 1
                    # 同段第二张及之后的图不重复挂载备注（避免一份备注被多张图误用）
                    note = note_for_first if bi == 0 else ''
                    saved = _save_image(part.blob, part.content_type, idx, note)
                    if saved is not None:
                        extracted.append(saved)
                # 段内若同时有文字，文字视为对该段图片的说明，不当作下一张图的备注
            elif para_text:
                # 纯文字段落：作为备注候选加入缓冲区
                pending_notes.append(para_text)

        if extracted:
            return extracted
        # python-docx 打开成功但没找到图片，可能是浮动图片或空文档，
        # 不立即返回，继续尝试回退策略以防遗漏
    except Exception as e:
        # WPS 生成的非标准 docx 缺 docProps/core.xml，python-docx 打开必失败；
        # 此为已知情况，静默走 zipfile 回退策略，不打印噪音
        if 'docProps/core.xml' not in str(e):
            print(f"   ℹ️ python-docx 解析失败，尝试直接解压提取：{str(e)[:80]}")

    # ---- 策略2：zipfile 直接解析 document.xml（不依赖 python-docx，仍可解析备注）----
    import zipfile
    import re as _re
    import xml.etree.ElementTree as ET
    extracted = []
    # OOXML 命名空间
    _W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    _A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    _R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    def _qn(ns, tag):
        return f"{{{ns}}}{tag}"

    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            # 读 rels：rId -> media 路径
            # 注意：WPS 生成的 docx 可能将同一 rId 定义多次（如 rId6 既是 image 又是 numbering），
            # 只记录 Type 以 /image 结尾的关系，且遇到同 rId 已有 image 记录时不覆盖
            rels_map = {}
            try:
                with z.open('word/_rels/document.xml.rels') as rf:
                    rels_root = ET.fromstring(rf.read())
                for rel in rels_root.iter():
                    rid = rel.get('Id')
                    target = rel.get('Target')
                    rtype = rel.get('Type', '')
                    if not (rid and target):
                        continue
                    # 只接受图片类型关系，非图片类型不记录
                    if not rtype.endswith('/image'):
                        continue
                    # 已有图片记录则不覆盖（保留第一次定义）
                    if rid not in rels_map:
                        rels_map[rid] = target  # 如 "media/image1.png"
            except Exception:
                pass

            # 读 document.xml，按段落顺序提取文字与图片
            with z.open('word/document.xml') as df:
                doc_root = ET.fromstring(df.read())
            body = doc_root.find(_qn(_W_NS, 'body'))
            if body is None:
                raise RuntimeError("document.xml 缺少 body 元素")

            idx = 0
            pending_notes = []
            for child in list(body):
                tag = child.tag.split('}')[-1]
                if tag != 'p':
                    # 非段落（如表格）：跳过，表格内图片暂不处理
                    continue
                # 段落文字
                texts = [t.text or '' for t in child.iter(_qn(_W_NS, 't'))]
                para_text = ''.join(texts).strip()
                # 段落内图片
                blips = list(child.iter(_qn(_A_NS, 'blip')))

                if blips:
                    note_for_first = '\n'.join(pending_notes)
                    pending_notes.clear()
                    for bi, blip in enumerate(blips):
                        rid = blip.get(_qn(_R_NS, 'embed'))
                        if not rid:
                            continue
                        media_rel = rels_map.get(rid)
                        if not media_rel:
                            continue
                        # rels 里 Target 是相对路径，如 "media/image1.png"，补 word/ 前缀
                        zip_path = 'word/' + media_rel if not media_rel.startswith('word/') else media_rel
                        ext = os.path.splitext(zip_path)[1].lower()
                        # 过滤非图片关系（numbering.xml/styles.xml 等）
                        if ext not in zip_ext_map:
                            continue
                        try:
                            blob = z.read(zip_path)
                        except Exception:
                            continue
                        content_type = zip_ext_map[ext]
                        idx += 1
                        note = note_for_first if bi == 0 else ''
                        saved = _save_image(blob, content_type, idx, note)
                        if saved is not None:
                            extracted.append(saved)
                elif para_text:
                    pending_notes.append(para_text)

    except Exception as e:
        raise RuntimeError(f"无法从文档提取图片：{e}")

    if not extracted:
        print(f"   ⚠️ 文档中未找到任何嵌入式图片：{os.path.basename(docx_path)}")
    return extracted


def batch_process(folder_path, docx_paths, dwg_path, root):
    """批量处理主程序入口：负责 docx 提取临时目录的创建与清理。

    图片来源支持两种：
    1. 一个文件夹内的所有图片
    2. 一个或多个 .docx 文档中嵌入的图片（按出现顺序提取）

    参数由主窗（_create_main_window）传入：
    folder_path: 图片文件夹路径或 None
    docx_paths: .docx 路径元组（可为空）
    dwg_path: 目标 DWG 文件路径
    root: 主窗口引用（用作 messagebox.parent）
    """
    temp_dir = tempfile.mkdtemp(prefix='autodraw_docx_') if docx_paths else None
    try:
        _batch_process_inner(folder_path, docx_paths, dwg_path, root, temp_dir)
    finally:
        # 正常返回或中途抛异常（含 OCR 引擎级故障）都要清掉临时目录，
        # 否则提取出的图片会在系统临时目录里不断累积
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


def _batch_process_inner(folder_path, docx_paths, dwg_path, root, temp_dir):
    """批处理主体。temp_dir 由 batch_process 创建，并在其 finally 中统一清理。"""
    print(f"AutoDraw 启动")
    print(f"图片来源: {'文件夹' if folder_path else 'Word文档'} | DWG: {os.path.basename(dwg_path)}")
    print("=" * 50)

    print("⏳ 正在加载 OCR 模型（首次运行会自动下载，请稍候）...")
    _get_ocr_engine()

    # 4. 收集所有待处理图片：[(完整路径, 显示文件名, 备注字符串), ...]
    #    先收集文件夹内图片，再提取 docx 内图片
    supported_formats = ('.png', '.jpg', '.jpeg', '.bmp')
    image_list = []   # [(完整路径, 显示文件名, 备注字符串), ...]

    if folder_path:
        try:
            filenames = sorted(os.listdir(folder_path))
        except OSError as e:
            # 文件夹被删除/无访问权限时给出可读原因，而不是让整轮处理崩在入口
            print(f"⚠️ 无法读取图片文件夹 {folder_path}: {e}")
            return
        for filename in filenames:
            if filename.lower().endswith(supported_formats):
                image_list.append((os.path.join(folder_path, filename), filename, ''))
        print(f"\n📁 文件夹图片：{len(image_list)} 张")

    if docx_paths:
        for docx_path in docx_paths:
            try:
                imgs = extract_images_from_docx(docx_path, temp_dir)
                image_list.extend(imgs)
                notes_count = sum(1 for _, _, n in imgs if n)
                print(f"📄 {os.path.basename(docx_path)}：提取到 {len(imgs)} 张图片（其中 {notes_count} 张含备注）")
            except Exception as e:
                print(f"⚠️ 读取 docx 失败 {os.path.basename(docx_path)}: {e}")

    if not image_list:
        print("\n⚠️ 未收集到任何图片，程序结束。")
        return

    print(f"\n🚀 共收集 {len(image_list)} 张图片，开始识别...\n{'=' * 50}")

    valid_points = []   # [(东坐标, 北坐标, 高程, 后缀, 文件名), ...]
    failures = []       # [(图片完整路径, 文件名, 失败原因), ...]
    total_images = len(image_list)
    _set_progress(0.0)
    for idx, (full_path, filename, note) in enumerate(image_list):
        result = extract_coords(full_path)

        # 5. 打印识别结果
        print(f"📄 文件: {result['文件名']}")
        print(f"   北坐标: {result['北坐标']}")
        print(f"   东坐标: {result['东坐标']}")
        print(f"   高  程: {result['高程']}")

        # 6. 收集识别成功的有效坐标（东坐标=X，北坐标=Y，高程用于文字标注）
        try:
            easting = float(result["东坐标"])
            northing = float(result["北坐标"])
            elevation = float(result["高程"])

            # 应用 docx 备注：解析 offset 与 suffix，修正高程 + 追加后缀
            # 备注与应用备注应紧邻显示（同一点的操作），最后才打印分隔线
            offset, suffix = 0.0, ''
            if note:
                parsed = _parse_elevation_note(note)
                if parsed:
                    offset, suffix = parsed
                    elevation_new = elevation + offset
                    print(f"   📝 备注: {note}")
                    print(f"   🔧 应用备注: 高程 {_trunc3(elevation)} + ({_trunc3(offset)}) = {_trunc3(elevation_new)}, 标注: {_trunc3(elevation_new)}{suffix}")
                    elevation = elevation_new
                else:
                    # 不可解析的备注（如纯"中顶"），仅显示备注本身
                    print(f"   📝 备注: {note}")
            valid_points.append((easting, northing, elevation, suffix, filename))
        except (ValueError, TypeError):
            # 识别失败或报错：记录原因，稍后统一归档
            east = str(result.get("东坐标", ""))
            reason = ("OCR识别异常: " + east) if east.startswith("报错") else "坐标识别失败（未提取到3个有效坐标值）"
            failures.append((full_path, filename, reason))
        # 分隔线统一在每张图片处理结束后打印（含备注应用情况）
        print("-" * 50)
        # OCR 阶段占总进度的 50%
        _set_progress((idx + 1) / total_images * 0.5)

    # 归档基础目录：有文件夹用文件夹；仅 docx 时用第一个 docx 所在目录，
    # 避免归档目录建在 temp_dir 内被清理时误删。
    if folder_path:
        archive_base = folder_path
    elif docx_paths:
        archive_base = os.path.dirname(os.path.abspath(docx_paths[0]))
    else:
        archive_base = None

    if not valid_points:
        print("\n⚠️ 没有识别到可用坐标，无需绘图。")
        _set_progress(1.0)
        archive_dir = archive_failed_images(archive_base, failures) if archive_base else None
        msg = f"共扫描 {total_images} 张图片，未识别到任何可用坐标。"
        if archive_dir:
            msg += f"\n失败图片已保存至:\n{archive_dir}"
        print(msg)
        _msg(messagebox.showwarning, "处理完成（无有效坐标）", msg, root)
        return

    # 7. 打开 DWG 并绘制红色圆
    print(f"\n📌 共识别到 {len(valid_points)} 组有效坐标。")
    print(f"🎯 正在 AutoCAD 中打开图纸并绘制圆(半径 {CIRCLE_RADIUS}m，红色)...\n{'=' * 50}")

    draw_failed = []
    cad_fatal = None
    try:
        count, draw_failed = draw_circles_on_dwg(dwg_path, valid_points)
    except Exception as e:
        cad_fatal = str(e)
        count = 0
        print(f"❌ AutoCAD 绘图失败: {e}")
        print("请确认本机已安装 AutoCAD，且 DWG 文件未被其他程序独占占用。")
        # CAD 整体不可用：所有待绘图片都按绘制失败归档
        # 原因不只可能是连接失败（也可能是图层名/文字样式预检不通过），故用中性描述
        draw_failed = [
            (pt[4], f"AutoCAD 绘图失败: {cad_fatal[:150]}")
            for pt in valid_points
        ]

    # 合并 OCR 失败与绘制失败的图片（按文件名去重）
    existing = {name for _, name, _ in failures}
    # 建立 文件名 -> 完整路径 的映射（含文件夹图片和 docx 临时图片）
    name_to_path = {fn: fp for fp, fn, _ in image_list}
    for failed_name, reason in draw_failed:
        if failed_name not in existing:
            failures.append((name_to_path.get(failed_name, failed_name), failed_name, reason))
            existing.add(failed_name)

    # 8. 归档失败图片
    archive_dir = archive_failed_images(archive_base, failures) if archive_base else None

    # 9. 弹窗汇总结果
    print("=" * 50)
    print(f"🎉 处理结束：共 {total_images} 张，成功 {count} 个，失败 {len(failures)} 个。")
    if archive_dir:
        print(f"📁 失败图片已保存至: {archive_dir}")
    print("📝 图纸保持打开状态，确认无误后请自行保存(Ctrl+S)。")

    if cad_fatal:
        msg = f"AutoCAD 绘图失败：{cad_fatal}\n\n失败图片已保存至:\n{archive_dir or archive_base}"
        _msg(messagebox.showerror, "处理异常", msg, root)
    else:
        msg = f"处理完成！\n\n共扫描图片：{total_images} 张\n成功绘制：{count} 个\n失败：{len(failures)} 个"
        if archive_dir:
            msg += f"\n\n失败图片已保存至:\n{archive_dir}"
        msg += "\n\n图纸保持打开，确认无误后请按 Ctrl+S 保存。"
        if failures:
            _msg(messagebox.showwarning, "处理完成（存在失败项）", msg, root)
        else:
            _msg(messagebox.showinfo, "处理完成", msg, root)

    # 清理 docx 临时提取目录的工作已在 batch_process 的 finally 中统一完成
    _set_progress(1.0)  # 处理完成，进度 100%


if __name__ == "__main__":
    # 设置任务栏归组标识：否则 windowed 模式打包后任务栏不显示程序图标
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass
    # 创建主集成窗体：上半部分文件选择 + 下半部分日志输出
    # 主窗关闭即程序退出，mainloop 阻塞至用户关闭窗口
    _ROOT, _ = _create_main_window()
    _ROOT.mainloop()