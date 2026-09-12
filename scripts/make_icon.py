"""生成应用图标：青绿底 + 白色文档页 + 矢量字母 A + 排版行，输出 icon.png 与多尺寸 icon.ico。

字母 A 用多边形路径绘制，不依赖系统字体，任何环境下输出一致。
用法：venv/Scripts/python scripts/make_icon.py
依赖：PySide6（绘制）、Pillow（ico 多尺寸封装）
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QRectF, Qt  # noqa: E402
from PySide6.QtGui import (QBrush, QColor, QGuiApplication, QImage,  # noqa: E402
                           QPainter, QPainterPath, QLinearGradient)

ASSET_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "app", "assets")

TEAL_DARK = "#0a5f57"
TEAL_MAIN = "#12897c"
TEAL_BRIGHT = "#12a594"
INK_LINE = "#c2cfce"


def letter_a_path(x, y, w, h):
    """几何无衬线字母 A：外轮廓挖去顶部三角孔，平顶、粗腿、实心横杠。"""
    path = QPainterPath()
    path.moveTo(x + 0.00 * w, y + 1.00 * h)
    path.lineTo(x + 0.33 * w, y + 0.00 * h)
    path.lineTo(x + 0.67 * w, y + 0.00 * h)
    path.lineTo(x + 1.00 * w, y + 1.00 * h)
    path.lineTo(x + 0.76 * w, y + 1.00 * h)
    path.lineTo(x + 0.69 * w, y + 0.79 * h)
    path.lineTo(x + 0.31 * w, y + 0.79 * h)
    path.lineTo(x + 0.24 * w, y + 1.00 * h)
    path.closeSubpath()
    hole = QPainterPath()
    hole.moveTo(x + 0.50 * w, y + 0.18 * h)
    hole.lineTo(x + 0.64 * w, y + 0.60 * h)
    hole.lineTo(x + 0.36 * w, y + 0.60 * h)
    hole.closeSubpath()
    return path.subtracted(hole)


def draw_icon(size=1024):
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    s = size / 1024.0

    # 圆角方形底：青绿渐变
    grad = QLinearGradient(0, 64 * s, 0, 960 * s)
    grad.setColorAt(0.0, QColor(TEAL_MAIN))
    grad.setColorAt(1.0, QColor(TEAL_DARK))
    bg = QPainterPath()
    bg.addRoundedRect(QRectF(64 * s, 64 * s, 896 * s, 896 * s), 210 * s, 210 * s)
    p.fillPath(bg, QBrush(grad))

    # 文档页（右上折角挖切）与柔和投影
    px, py, pw, ph = 292 * s, 218 * s, 452 * s, 588 * s
    fold = 104 * s
    shadow = QPainterPath()
    shadow.addRoundedRect(QRectF(px + 14 * s, py + 20 * s, pw, ph), 32 * s, 32 * s)
    p.fillPath(shadow, QColor(20, 60, 56, 70))
    page = QPainterPath()
    page.addRoundedRect(QRectF(px, py, pw, ph), 32 * s, 32 * s)
    cut = QPainterPath()
    cut.moveTo(px + pw - fold, py)
    cut.lineTo(px + pw, py)
    cut.lineTo(px + pw, py + fold)
    cut.closeSubpath()
    p.fillPath(page.subtracted(cut), QBrush(QColor("#ffffff")))
    fold_path = QPainterPath()
    fold_path.moveTo(px + pw - fold, py)
    fold_path.lineTo(px + pw - fold, py + fold)
    fold_path.lineTo(px + pw, py + fold)
    fold_path.closeSubpath()
    p.fillPath(fold_path, QBrush(QColor("#d7e5e3")))

    # 矢量字母 A（排版主题）
    p.fillPath(letter_a_path(px + 56 * s, py + 52 * s, 330 * s, 330 * s),
               QBrush(QColor(TEAL_DARK)))

    # 文本行：第一条青绿色（已套用格式），其余灰色
    for ly, lw, color in ((py + 424 * s, pw - 128 * s, TEAL_BRIGHT),
                          (py + 484 * s, pw - 210 * s, INK_LINE),
                          (py + 544 * s, pw - 128 * s, INK_LINE)):
        bar = QPainterPath()
        bar.addRoundedRect(QRectF(px + 64 * s, ly, lw, 32 * s), 16 * s, 16 * s)
        p.fillPath(bar, QBrush(QColor(color)))
    p.end()
    return img


def main():
    os.makedirs(ASSET_DIR, exist_ok=True)
    QGuiApplication(sys.argv)
    img = draw_icon()
    png_path = os.path.join(ASSET_DIR, "icon.png")
    img.save(png_path)
    print("saved:", png_path)

    from PIL import Image
    ico_path = os.path.join(ASSET_DIR, "icon.ico")
    Image.open(png_path).save(
        ico_path, sizes=[(256, 256), (128, 128), (64, 64), (48, 48),
                         (32, 32), (24, 24), (16, 16)])
    print("saved:", ico_path)


if __name__ == "__main__":
    main()
