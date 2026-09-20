# -*- coding: utf-8 -*-
import cv2
import numpy as np
import re
import os
import shutil
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
                "东坐标": f"{float(best_numbers[-3]):.3f}",
                "北坐标": f"{float(best_numbers[-2]):.3f}",
                "高程": f"{float(best_numbers[-1]):.3f}"
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
            text_obj = model_space.AddText(f"{elevation:.3f}", center, TEXT_HEIGHT)
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


def batch_process():
    """批量处理主程序：OCR 识别坐标 -> 在指定 DWG 上绘制红色圆"""
    # 1. 弹出文件夹选择窗口
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title="请选择包含坐标照片的文件夹")

    if not folder_path:
        print("未选择文件夹，程序退出。")
        return

    # 2. 立即选择要绘制的 DWG 文件（在 OCR 开始前选好，无需等待识别完成）
    dwg_path = filedialog.askopenfilename(
        title="请选择要在上面绘制圆的 .dwg 文件",
        filetypes=[("DWG 文件", "*.dwg"), ("所有文件", "*.*")]
    )
    if not dwg_path:
        print("未选择 DWG 文件，程序退出。")
        return

    print("⏳ 正在加载 OCR 模型（首次运行会自动下载，请稍候）...")
    _get_ocr_engine()

    # 3. 遍历文件夹中的图片并识别坐标（识别期间无需人工等待操作）
    supported_formats = ('.png', '.jpg', '.jpeg', '.bmp')
    print(f"\n🚀 正在处理文件夹: {folder_path}\n{'=' * 50}")

    valid_points = []   # [(东坐标, 北坐标, 高程, 文件名), ...]
    failures = []       # [(图片完整路径, 文件名, 失败原因), ...]
    total_images = 0
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(supported_formats):
            total_images += 1
            full_path = os.path.join(folder_path, filename)
            result = extract_coords(full_path)

            # 4. 打印识别结果
            print(f"📄 文件: {result['文件名']}")
            print(f"   北坐标: {result['北坐标']}")
            print(f"   东坐标: {result['东坐标']}")
            print(f"   高  程: {result['高程']}")
            print("-" * 50)

            # 5. 收集识别成功的有效坐标（东坐标=X，北坐标=Y，高程用于文字标注）
            try:
                easting = float(result["东坐标"])
                northing = float(result["北坐标"])
                elevation = float(result["高程"])
                valid_points.append((easting, northing, elevation, result["文件名"]))
            except (ValueError, TypeError):
                # 识别失败或报错：记录原因，稍后统一归档
                east = str(result.get("东坐标", ""))
                reason = ("OCR识别异常: " + east) if east.startswith("报错") else "坐标识别失败（未提取到3个有效坐标值）"
                failures.append((full_path, filename, reason))

    if not valid_points:
        print("\n⚠️ 没有识别到可用坐标，无需绘图。")
        archive_dir = archive_failed_images(folder_path, failures)
        msg = f"共扫描 {total_images} 张图片，未识别到任何可用坐标。"
        if archive_dir:
            msg += f"\n失败图片已保存至:\n{archive_dir}"
        print(msg)
        messagebox.showwarning("处理完成（无有效坐标）", msg, parent=root)
        return

    # 6. 打开 DWG 并绘制红色圆
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
    name_to_path = {pt[3]: os.path.join(folder_path, pt[3]) for pt in valid_points}
    for failed_name, reason in draw_failed:
        if failed_name not in existing:
            failures.append((name_to_path.get(failed_name, failed_name), failed_name, reason))
            existing.add(failed_name)

    # 7. 归档失败图片
    archive_dir = archive_failed_images(folder_path, failures)

    # 8. 弹窗汇总结果
    print("=" * 50)
    print(f"🎉 处理结束：共 {total_images} 张，成功 {count} 个，失败 {len(failures)} 个。")
    if archive_dir:
        print(f"📁 失败图片已保存至: {archive_dir}")
    print("📝 图纸保持打开状态，确认无误后请自行保存(Ctrl+S)。")

    if cad_fatal:
        msg = f"AutoCAD 绘图失败：{cad_fatal}\n\n失败图片已保存至:\n{archive_dir or folder_path}"
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


if __name__ == "__main__":
    batch_process()