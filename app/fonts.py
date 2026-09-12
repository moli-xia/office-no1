"""预设所需字体的检测与安装。

随程序附带的字体放在 Fonts/ 目录（打包进 exe 后位于临时解压目录）。
安装采用 Windows 10 1809+ 的“为当前用户安装”方式：复制到
%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts 并写入 HKCU 注册表，不需要管理员权限。
"""

import ctypes
import os
import shutil
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_DIR = os.path.join(ROOT_DIR, "Fonts")

# names[0] 为中文名（也是预设里使用的名称），其余为字体文件内的英文名
BUNDLED_FONTS = [
    {"file": "仿宋_GB2312.ttf", "names": ["仿宋_GB2312", "FangSong_GB2312"], "usage": "公文正文"},
    {"file": "楷体_GB2312.ttf", "names": ["楷体_GB2312", "KaiTi_GB2312"], "usage": "公文二级标题"},
    {"file": "方正小标宋简.TTF", "names": ["方正小标宋简体", "FZXiaoBiaoSong-B05S"], "usage": "公文标题"},
]

_REG_FONTS = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"


def _registered_font_names():
    """读取 HKLM / HKCU 中已登记的字体名（小写、去掉“(TrueType)”后缀）。"""
    names = set()
    if not sys.platform.startswith("win"):
        return names
    import winreg
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(hive, _REG_FONTS)
        except OSError:
            continue
        with key:
            i = 0
            while True:
                try:
                    name, _value, _type = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                names.add(name.split(" (")[0].strip().lower())
    return names


def _qt_families():
    try:
        from PySide6.QtGui import QFontDatabase
        from PySide6.QtWidgets import QApplication
        if QApplication.instance() is None:
            return set()
        return {f.lower() for f in QFontDatabase.families()}
    except Exception:
        return set()


def is_installed(font, registered=None, families=None):
    registered = _registered_font_names() if registered is None else registered
    families = _qt_families() if families is None else families
    for name in font["names"]:
        n = name.lower()
        if n in registered or n in families:
            return True
    return False


def bundled_path(font):
    return os.path.join(FONT_DIR, font["file"])


def missing_fonts():
    """返回本机未安装、且程序附带了字体文件的字体列表。"""
    registered = _registered_font_names()
    families = _qt_families()
    return [f for f in BUNDLED_FONTS
            if os.path.isfile(bundled_path(f)) and not is_installed(f, registered, families)]


def user_font_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~\\AppData\\Local")
    return os.path.join(base, "Microsoft", "Windows", "Fonts")


def install_font(font):
    """为当前用户安装一款字体。成功返回 None，失败返回错误说明。"""
    src = bundled_path(font)
    if not os.path.isfile(src):
        return f"找不到字体文件：{src}"
    if not sys.platform.startswith("win"):
        return "仅支持 Windows"
    try:
        import winreg
        folder = user_font_dir()
        os.makedirs(folder, exist_ok=True)
        dst = os.path.join(folder, font["file"])
        if not (os.path.isfile(dst) and os.path.getsize(dst) == os.path.getsize(src)):
            if os.path.exists(dst):
                os.chmod(dst, 0o666)
            shutil.copyfile(src, dst)
            os.chmod(dst, 0o666)  # 附带的字体文件可能是只读的，别把只读属性一起带过去
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_FONTS) as key:
            winreg.SetValueEx(key, f"{font['names'][0]} (TrueType)", 0, winreg.REG_SZ, dst)
        # 让当前会话立即可用，并通知其它程序字体列表已变化
        ctypes.windll.gdi32.AddFontResourceW(dst)
        HWND_BROADCAST, WM_FONTCHANGE, SMTO_ABORTIFHUNG = 0xFFFF, 0x001D, 0x0002
        ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0,
                                                 SMTO_ABORTIFHUNG, 1000, None)
    except Exception as error:
        return f"{type(error).__name__}: {error}"
    return None


def install_fonts(fonts):
    """批量安装，返回 (成功列表, [(字体, 错误)])。"""
    ok, failed = [], []
    for f in fonts:
        err = install_font(f)
        (failed.append((f, err)) if err else ok.append(f))
    return ok, failed
