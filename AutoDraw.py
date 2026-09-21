# -*- coding: utf-8 -*-
import cv2
import numpy as np
import re
import os
import shutil
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import win32com.client
import pythoncom

# RapidOCR 引擎单例：模型只加载一次。首次运行会自动下载 ONNX 模型。
_OCR_ENGINE = None

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


def _trunc3(x):
    """截断到小数点后 3 位（不四舍五入），返回字符串。

    用字符串切片而非 round()，避免 343.4956 被四舍五入成 343.496。
    取 10 位小数再切片，规避浮点数精度导致的截断错位。
    """
    s = f"{float(x):.10f}"
    dot = s.index('.')
    return s[:dot + 4]


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
    return _normalize_ocr_text("\n".join(_reading_order_texts(result)))


def _coords_from_text(text):
    all_numbers = re.findall(r"-?\d+\.\d+", text)
    return [n for n in all_numbers if abs(float(n)) > 100]


def extract_coords(image_path):
    """从单张图片中提取坐标。

    仪器屏幕的数值排列顺序为：东坐标 -> 北坐标 -> 高程，
    因此取筛选后最后三个数依次作为 东坐标、北坐标、高程。
    """
    try:
        img = _load_bgr(image_path)
        if img is None:
            return {"文件名": os.path.basename(image_path), "东坐标": "识别失败",
                    "北坐标": "图片无法读取或已损坏", "高程": "识别失败"}

        best_numbers = []
        for candidate in (img, _enhance_for_blur(img)):
            valid_coords = _coords_from_text(_ocr_to_text(candidate))
            if len(valid_coords) >= 3:
                best_numbers = valid_coords
                break
            if len(valid_coords) > len(best_numbers):
                best_numbers = valid_coords

        if len(best_numbers) >= 3:
            return {
                "文件名": os.path.basename(image_path),
                "东坐标": _trunc3(best_numbers[-3]),
                "北坐标": _trunc3(best_numbers[-2]),
                "高程": _trunc3(best_numbers[-1])
            }
        return {"文件名": os.path.basename(image_path), "东坐标": "识别失败", "北坐标": "识别失败",
                "高程": "识别失败"}

    except Exception as e:
        return {"文件名": os.path.basename(image_path), "东坐标": f"报错: {e}", "北坐标": "",
                "高程": ""}


def _ensure_layer(doc, layer_name):
    """获取指定图层，不存在则自动创建"""
    try:
        return doc.Layers.Item(layer_name)
    except Exception:
        return doc.Layers.Add(layer_name)


def _draw_on_doc(doc, points):
    """在给定文档上绘制圆与高程文字标注。

    points: [(东坐标X, 北坐标Y, 高程, 文件名), ...]
    返回 (成功数量, 失败列表[(文件名, 失败原因)])。
    单个点绘制失败不影响其他点；若圆已画出但文字失败，会删除该圆保持整组回滚。
    """
    _ensure_layer(doc, LAYER_CIRCLE)
    _ensure_layer(doc, LAYER_TEXT)

    model_space = doc.ModelSpace
    drawn = 0
    failed = []
    for easting, northing, elevation, source_file in points:
        circle = None
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
            text_obj = model_space.AddText(_trunc3(elevation), center, TEXT_HEIGHT)
            text_obj.Layer = LAYER_TEXT
            text_obj.color = TEXT_COLOR      # 绿色
            text_obj.StyleName = TEXT_STYLE
            text_obj.ScaleFactor = TEXT_WIDTH_FACTOR
            text_obj.HorizontalAlignment = 4  # acHorizontalAlignmentMiddle：以对齐点为文字中心
            text_obj.TextAlignmentPoint = center
            text_obj.Update()

            drawn += 1
            print(f"   ✅ 已绘制圆+高程标注: X(东)={easting}  Y(北)={northing}  高程={elevation}  <- {source_file}")
        except Exception as e:
            # 圆已画出但后续失败：删除半成品圆，避免图上留下没有标注的圆
            if circle is not None:
                try:
                    circle.Delete()
                except Exception:
                    pass
            reason = str(e).replace('\n', ' ')[:200]
            failed.append((source_file, f"绘制失败: {reason}"))
            print(f"   ❌ 绘制失败: {source_file} -> {reason}")

    return drawn, failed


def draw_circles_on_dwg(dwg_path, points):
    """打开指定 DWG 文件，在每个坐标点绘制红色圆并标注高程文字。

    points: [(东坐标X, 北坐标Y, 高程, 文件名), ...]
    坐标系：东坐标 -> CAD 的 X 轴，北坐标 -> CAD 的 Y 轴
    """
    abs_dwg = os.path.abspath(dwg_path)

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
            time.sleep(1)
        if doc is None:
            for open_doc in acad.Documents:
                if os.path.abspath(open_doc.FullName).lower() == abs_dwg.lower():
                    doc = open_doc
                    break
    doc.Activate()

    drawn, failed = _draw_on_doc(doc, points)

    # 视图缩放到本次成功绘制的区域（外扩 ZOOM_MARGIN 米），而不是整张图，
    # 避免 ZoomExtents 把大坐标图纸缩得过小、难以确认刚绘制的点
    success_points = [pt for pt in points
                      if pt[3] not in {name for name, _ in failed}]
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
        acad.ZoomWindow(lower, upper)
    return drawn, failed


def archive_failed_images(folder_path, failures):
    """把失败图片复制到照片目录下带时间戳的新文件夹，并写入失败原因清单。

    failures: [(图片完整路径, 文件名, 失败原因), ...]
    返回归档文件夹路径（无失败项时返回 None）。
    """
    if not failures:
        return None

    archive_dir = os.path.join(
        folder_path, f"识别失败图片_{datetime.now():%Y%m%d_%H%M%S}"
    )
    os.makedirs(archive_dir, exist_ok=True)

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
    log_path = os.path.join(archive_dir, "失败原因.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        f.write(f"失败数量: {len(failures)}\n")
        f.write("=" * 60 + "\n")
        for _, filename, reason in failures:
            f.write(f"{filename}\t{reason}\n")

    return archive_dir


def extract_images_from_docx(docx_path, output_dir):
    """从 docx 文档中按出现顺序提取所有嵌入式图片到 output_dir。

    返回 [(保存后的完整路径, 显示用文件名), ...]。
    文件名格式：{文档名(去扩展名)}_{序号:03d}.{扩展名}，便于追溯来源。

    解析策略：
    1. 优先用 python-docx 解析（能保证图片在文档中的出现顺序）；
    2. 若文档结构不标准（如 WPS 生成的 docx 缺少 docProps/core.xml 导致
       python-docx 无法打开），则回退到直接用 zipfile 解压 word/media/
       目录，按文件名序号排序提取。
    """
    from docx import Document
    from docx.enum.shape import WD_INLINE_SHAPE

    doc_base = os.path.splitext(os.path.basename(docx_path))[0]

    # 内容类型到扩展名的映射
    ext_map = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/jpg': '.jpg',
        'image/gif': '.gif',
        'image/bmp': '.bmp',
        'image/tiff': '.tiff',
        'image/x-emf': '.emf',
        'image/x-wmf': '.wmf',
    }
    # zip 内扩展名到 content_type 的反向映射（回退模式用）
    zip_ext_map = {
        '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
        '.gif': 'image/gif', '.bmp': 'image/bmp', '.tiff': 'image/tiff',
        '.emf': 'image/x-emf', '.wmf': 'image/x-wmf',
    }

    def _save_image(blob, content_type, idx):
        ext = ext_map.get(content_type, '.png')
        filename = f"{doc_base}_{idx:03d}{ext}"
        save_path = os.path.join(output_dir, filename)
        with open(save_path, 'wb') as f:
            f.write(blob)
        return (save_path, filename)

    # ---- 策略1：python-docx（顺序可靠）----
    try:
        doc = Document(docx_path)
        extracted = []
        idx = 0
        for shape in doc.inline_shapes:
            # 只处理图片类型（忽略公式、图表等）
            if shape.type != WD_INLINE_SHAPE.PICTURE:
                continue
            try:
                image_part = shape._inline.graphic.graphicData.pic.blipFill.blip
                rId = image_part.embed
                image_part = doc.part.related_parts[rId]
            except Exception:
                try:
                    image_part = shape.image_part
                except Exception:
                    continue
            idx += 1
            extracted.append(_save_image(image_part.blob, image_part.content_type, idx))
        if extracted:
            return extracted
        # python-docx 打开成功但没找到图片，可能是浮动图片或空文档，
        # 不立即返回，继续尝试回退策略以防遗漏
    except Exception as e:
        print(f"   ℹ️ python-docx 解析失败，尝试直接解压提取：{str(e)[:80]}")

    # ---- 策略2：zipfile 直接解压 word/media/（兼容性优先）----
    import zipfile
    import re as _re
    extracted = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            media_names = [n for n in z.namelist()
                           if n.startswith('word/media/') and not n.endswith('/')]
            # 按文件名中的数字序号排序，尽量还原插入顺序
            def _sort_key(name):
                m = _re.search(r'(\d+)', os.path.basename(name))
                return (int(m.group(1)) if m else 9999, name)
            media_names.sort(key=_sort_key)

            idx = 0
            for name in media_names:
                ext = os.path.splitext(name)[1].lower()
                if ext not in zip_ext_map:
                    continue
                idx += 1
                blob = z.read(name)
                extracted.append(_save_image(blob, zip_ext_map[ext], idx))
    except Exception as e:
        raise RuntimeError(f"无法从文档提取图片：{e}")

    if not extracted:
        print(f"   ⚠️ 文档中未找到任何嵌入式图片：{os.path.basename(docx_path)}")
    return extracted


def _ask_source_type(root):
    """弹窗让用户选择图片来源类型，返回 'folder' 或 'docx'，取消返回 None。"""
    choice = {'value': None}
    win = tk.Toplevel(root)
    win.title("选择图片来源")
    win.attributes('-topmost', True)
    win.resizable(False, False)

    tk.Label(win, text="请选择坐标照片的来源：", font=("Microsoft YaHei", 11),
             padx=30, pady=20).pack()

    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=(0, 20))

    def pick_folder():
        choice['value'] = 'folder'
        win.destroy()

    def pick_docx():
        choice['value'] = 'docx'
        win.destroy()

    tk.Button(btn_frame, text="📁 图片文件夹", font=("Microsoft YaHei", 11),
              width=16, height=2, command=pick_folder).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="📄 Word 文档(.docx)", font=("Microsoft YaHei", 11),
              width=16, height=2, command=pick_docx).pack(side=tk.LEFT, padx=10)

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.update_idletasks()
    # 居中
    x = (win.winfo_screenwidth() - win.winfo_reqwidth()) // 2
    y = (win.winfo_screenheight() - win.winfo_reqheight()) // 2
    win.geometry(f"+{x}+{y}")
    win.grab_set()
    root.wait_window(win)
    return choice['value']


def batch_process():
    """批量处理主程序：OCR 识别坐标 -> 在指定 DWG 上绘制红色圆

    图片来源支持两种：
    1. 一个文件夹内的所有图片
    2. 一个或多个 .docx 文档中嵌入的图片（按出现顺序提取）
    运行后先弹窗让用户选择来源类型，再进行后续选择。
    """
    # 1. 弹窗选择图片来源类型
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    source_type = _ask_source_type(root)
    if source_type is None:
        print("未选择图片来源类型，程序退出。")
        return

    folder_path = None
    docx_paths = ()

    # 2. 根据来源类型弹出对应的选择框
    if source_type == 'folder':
        folder_path = filedialog.askdirectory(title="请选择包含坐标照片的文件夹")
        if not folder_path:
            print("未选择文件夹，程序退出。")
            return
    else:  # docx
        docx_paths = filedialog.askopenfilenames(
            title="请选择包含坐标照片的 .docx 文档（可多选）",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if not docx_paths:
            print("未选择 docx 文档，程序退出。")
            return

    # 3. 立即选择要绘制的 DWG 文件（在 OCR 开始前选好，无需等待识别完成）
    dwg_path = filedialog.askopenfilename(
        title="请选择要在上面绘制圆的 .dwg 文件",
        filetypes=[("DWG 文件", "*.dwg"), ("所有文件", "*.*")]
    )
    if not dwg_path:
        print("未选择 DWG 文件，程序退出。")
        return

    print("⏳ 正在加载 OCR 模型（首次运行会自动下载，请稍候）...")
    _get_ocr_engine()

    # 4. 收集所有待处理图片：[(完整路径, 显示文件名), ...]
    #    先收集文件夹内图片，再提取 docx 内图片
    supported_formats = ('.png', '.jpg', '.jpeg', '.bmp')
    image_list = []   # [(完整路径, 显示文件名), ...]
    temp_dir = None   # docx 提取图片的临时目录，结束时清理

    if folder_path:
        for filename in sorted(os.listdir(folder_path)):
            if filename.lower().endswith(supported_formats):
                image_list.append((os.path.join(folder_path, filename), filename))
        print(f"\n📁 文件夹图片：{len(image_list)} 张")

    if docx_paths:
        temp_dir = tempfile.mkdtemp(prefix='autodraw_docx_')
        for docx_path in docx_paths:
            try:
                imgs = extract_images_from_docx(docx_path, temp_dir)
                image_list.extend(imgs)
                print(f"📄 {os.path.basename(docx_path)}：提取到 {len(imgs)} 张图片")
            except Exception as e:
                print(f"⚠️ 读取 docx 失败 {os.path.basename(docx_path)}: {e}")

    if not image_list:
        print("\n⚠️ 未收集到任何图片，程序结束。")
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
        return

    print(f"\n🚀 共收集 {len(image_list)} 张图片，开始识别...\n{'=' * 50}")

    valid_points = []   # [(东坐标, 北坐标, 高程, 文件名), ...]
    failures = []       # [(图片完整路径, 文件名, 失败原因), ...]
    total_images = len(image_list)
    for full_path, filename in image_list:
        result = extract_coords(full_path)

        # 5. 打印识别结果
        print(f"📄 文件: {result['文件名']}")
        print(f"   北坐标: {result['北坐标']}")
        print(f"   东坐标: {result['东坐标']}")
        print(f"   高  程: {result['高程']}")
        print("-" * 50)

        # 6. 收集识别成功的有效坐标（东坐标=X，北坐标=Y，高程用于文字标注）
        try:
            easting = float(result["东坐标"])
            northing = float(result["北坐标"])
            elevation = float(result["高程"])
            valid_points.append((easting, northing, elevation, filename))
        except (ValueError, TypeError):
            # 识别失败或报错：记录原因，稍后统一归档
            east = str(result.get("东坐标", ""))
            reason = ("OCR识别异常: " + east) if east.startswith("报错") else "坐标识别失败（未提取到3个有效坐标值）"
            failures.append((full_path, filename, reason))

    # 归档基础目录：优先用文件夹路径（有文件夹时），否则用临时目录（仅 docx 时）
    archive_base = folder_path or temp_dir

    if not valid_points:
        print("\n⚠️ 没有识别到可用坐标，无需绘图。")
        archive_dir = archive_failed_images(archive_base, failures) if archive_base else None
        msg = f"共扫描 {total_images} 张图片，未识别到任何可用坐标。"
        if archive_dir:
            msg += f"\n失败图片已保存至:\n{archive_dir}"
        print(msg)
        messagebox.showwarning("处理完成（无有效坐标）", msg, parent=root)
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
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
        draw_failed = [
            (pt[3], f"AutoCAD连接/打开图纸失败: {cad_fatal[:150]}")
            for pt in valid_points
        ]

    # 合并 OCR 失败与绘制失败的图片（按文件名去重）
    existing = {name for _, name, _ in failures}
    # 建立 文件名 -> 完整路径 的映射（含文件夹图片和 docx 临时图片）
    name_to_path = {fn: fp for fp, fn in image_list}
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
        messagebox.showerror("处理异常", msg, parent=root)
    else:
        msg = f"处理完成！\n\n共扫描图片：{total_images} 张\n成功绘制：{count} 个\n失败：{len(failures)} 个"
        if archive_dir:
            msg += f"\n\n失败图片已保存至:\n{archive_dir}"
        msg += "\n\n图纸保持打开，确认无误后请按 Ctrl+S 保存。"
        if failures:
            messagebox.showwarning("处理完成（存在失败项）", msg, parent=root)
        else:
            messagebox.showinfo("处理完成", msg, parent=root)

    # 清理 docx 临时提取目录（失败图片已复制到归档目录，原临时文件可删）
    if temp_dir:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    batch_process()