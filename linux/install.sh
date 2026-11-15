#!/usr/bin/env bash
# ────────────────────────────────────────────────────────────────
# 材料排版工具 · 国产 Linux 一键安装（deepin / 统信UOS / 麒麟Kylin）
# 用法：
#   ./install.sh            完整安装（打包独立可执行程序 + 桌面图标）
#   ./install.sh --no-build 免打包安装（Python 脚本直跑，功能相同）
# 可重复运行（幂等）：自动清理并重建运行环境。
# 本目录自包含：app.py / gongwen.py / requirements.txt / icon.svg 齐备。
# ────────────────────────────────────────────────────────────────
set -uo pipefail

APP_NAME="材料排版工具"
INSTALL_DIR="$HOME/.local/opt/clpb"
DESKTOP_FILE="$HOME/.local/share/applications/clpb.desktop"
MIRROR="https://pypi.tuna.tsinghua.edu.cn/simple"
NO_BUILD=0
[[ "${1:-}" == "--no-build" ]] && NO_BUILD=1

step() { printf '\033[32m▶ %s\033[0m\n' "$1"; }
warn() { printf '\033[33m· %s\033[0m\n' "$1"; }
die()  { printf '\033[31m✘ %s\033[0m\n' "$1" >&2; exit 1; }

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# 源码目录：优先本目录（自包含）；兼容完整项目布局（linux/ 的上一级）
SRC_DIR="$SCRIPT_DIR"
[ -f "$SRC_DIR/app.py" ] || SRC_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
{ [ -f "$SRC_DIR/app.py" ] && [ -f "$SRC_DIR/gongwen.py" ]; } \
  || die "未找到 app.py / gongwen.py：请保持 linux/ 目录内容完整后重试"

SUDO=""
if [ "$(id -u)" -ne 0 ]; then
  command -v sudo >/dev/null 2>&1 || die "当前非 root 且未安装 sudo，请以 root 运行"
  SUDO="sudo"
fi

apt_install() {   # apt_install 包名…（已安装则跳过）
  local missing=() p
  for p in "$@"; do
    dpkg -s "$p" >/dev/null 2>&1 || missing+=("$p")
  done
  [ ${#missing[@]} -eq 0 ] && return 0
  warn "安装系统组件：${missing[*]}"
  $SUDO apt-get update -qq >/dev/null 2>&1 || true
  $SUDO apt-get install -y "${missing[@]}"
}

echo "════════ $APP_NAME · 安装到 $INSTALL_DIR ════════"

# ── 1. Python ──
step "检查 Python 环境…"
command -v python3 >/dev/null 2>&1 || die "未找到 python3，请执行：${SUDO} apt install python3"
python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3,7) else 1)' \
  || die "需要 Python ≥ 3.7（当前 $(python3 -V 2>&1)）"
PY_VER="$(python3 -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
echo "  · Python $PY_VER ($(python3 -V 2>&1))"
PY_OLD=0
python3 -c 'import sys; sys.exit(0 if sys.version_info < (3,8) else 1)' && PY_OLD=1

# ── 2. tkinter ──
python3 -c 'import tkinter' >/dev/null 2>&1 \
  || { apt_install python3-tk || die "安装 python3-tk 失败，请手动执行：${SUDO} apt install python3-tk"; }

# ── 3. 运行环境：venv 优先，失败自动降级为 系统 pip --user ──
step "准备运行环境…"
mkdir -p "$INSTALL_DIR/app" "$INSTALL_DIR/bin" "$INSTALL_DIR/share"

set_run_mode_venv() {
  PY="$INSTALL_DIR/venv/bin/python"
  PIP_CMD=("$PY" -m pip)
  PIP_USER=()
}
set_run_mode_system() {
  PY="$(command -v python3)"
  PIP_CMD=(python3 -m pip)
  PIP_USER=(--user)
}
make_venv() {   # $1 = 额外参数（可选 --system-site-packages）
  rm -rf "$INSTALL_DIR/venv"
  python3 -m venv ${1:-} "$INSTALL_DIR/venv" >/dev/null 2>&1 \
    && [ -x "$INSTALL_DIR/venv/bin/pip" ]    # 必须真正带 pip 才算成功
}
VENV_OK=0
if make_venv; then
  VENV_OK=1
else
  warn "创建 venv 失败（系统缺 python3-venv），尝试补装…"
  apt_install python3-venv >/dev/null 2>&1 || true
  if make_venv; then
    VENV_OK=1
  else
    apt_install "python${PY_VER}-venv" >/dev/null 2>&1 || true
    make_venv && VENV_OK=1
  fi
fi
if [ "$VENV_OK" -eq 1 ]; then
  set_run_mode_venv
  echo "  · 独立运行环境（venv）就绪"
else
  warn "本机 venv 不可用，改用系统 pip --user 模式（功能相同）"
  python3 -m pip --version >/dev/null 2>&1 \
    || { apt_install python3-pip || die "无法安装 pip，请手动执行：${SUDO} apt install python3-pip"; }
  set_run_mode_system
fi

# pip 安装：先静默走国内镜像，失败则回退官方源并显示完整日志（便于定位问题）
pip_install() {   # pip_install 包…；失败返回非零（不直接退出）
  "${PIP_CMD[@]}" install -q ${PIP_USER[@]+"${PIP_USER[@]}"} "$@" -i "$MIRROR" >/dev/null 2>&1 \
    && return 0
  warn "安装重试（回退官方源并显示日志）：$*"
  "${PIP_CMD[@]}" install ${PIP_USER[@]+"${PIP_USER[@]}"} "$@"
}

# ── 4. 升级构建工具（关键：Python3.7 venv 自带 pip 18.1 过老，
#        无法解析新版包元数据 pyproject.toml 与 manylinux2014 wheel 标签）──
step "升级 pip / setuptools / wheel…"
if [ "$PY_OLD" -eq 1 ]; then
  BUILD_TOOLS=("pip==24.0" "setuptools==68.0.0" "wheel==0.42.0")   # 支持 3.7 的最后版本
else
  BUILD_TOOLS=("pip" "setuptools" "wheel")
fi
pip_install "${BUILD_TOOLS[@]}" || warn "构建工具升级未成功，沿用现有版本继续尝试…"

# ── 5. 核心依赖 python-docx（失败走 apt 兜底）──
step "安装依赖（首次需联网）…"
if ! pip_install -r "$SCRIPT_DIR/requirements.txt"; then
  warn "pip 安装 python-docx 失败，尝试系统源兜底（apt install python3-docx）…"
  apt_install python3-docx \
    || die "python-docx 安装失败：请检查网络，或手动执行 ${SUDO} apt install python3-docx 后重跑本脚本"
  if [ "$VENV_OK" -eq 1 ]; then
    # 重建 venv 并放行系统包，使 apt 版 python-docx 可见
    rm -rf "$INSTALL_DIR/venv"
    if make_venv "--system-site-packages"; then
      set_run_mode_venv
    else
      VENV_OK=0
      set_run_mode_system
    fi
  fi
fi

# ── 6. 可选依赖：拖拽组件（失败不影响使用，自动降级为按钮选择文件）──
pip_install "tkinterdnd2>=0.4.0" \
  || warn "拖拽组件未装成：不影响使用，用「浏览…」按钮选择文件即可"

"$PY" -c 'import docx' >/dev/null 2>&1 \
  || die "python-docx 仍不可用，请把上方完整报错日志反馈给开发人员"
if "$PY" -c 'import tkinterdnd2' >/dev/null 2>&1; then
  echo "  · 拖拽组件：就绪"
else
  warn "拖拽组件：不可用（可用「浏览…」按钮选择文件，其余功能完整）"
fi

# ── 6.5 转换引擎：旧格式 .doc/.wps 需要 LibreOffice（WPS Linux 版不支持后台转换）──
step "检查文档转换引擎…"
if command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1 \
   || [ -x /usr/lib/libreoffice/program/soffice ]; then
  echo "  · LibreOffice：已就绪（.doc/.wps 可自动转换）"
else
  warn "未检测到 LibreOffice，尝试安装（用于 .doc/.wps 自动转换；.docx 不受影响）…"
  if $SUDO apt-get install -y --no-install-recommends libreoffice-writer >/dev/null 2>&1 \
     && command -v soffice >/dev/null 2>&1; then
    echo "  · LibreOffice：安装就绪"
  else
    warn "LibreOffice 未装成：旧格式 .doc/.wps 需先在 WPS 中另存为 .docx（.docx 排版不受影响）；"
    warn "  稍后可手动安装：${SUDO} apt install libreoffice-writer"
  fi
fi

# ── 7. 程序文件 ──
step "复制程序文件…"
cp -f "$SRC_DIR/app.py" "$SRC_DIR/gongwen.py" "$INSTALL_DIR/app/" || die "复制程序文件失败"
[ -f "$SCRIPT_DIR/icon.svg" ] && cp -f "$SCRIPT_DIR/icon.svg" "$INSTALL_DIR/share/"

# ── 8. 打包独立可执行程序（可选）──
BIN="$INSTALL_DIR/bin/$APP_NAME"
if [ "$NO_BUILD" -eq 0 ]; then
  step "打包独立可执行程序（PyInstaller，本机原生构建）…"
  if [ "$PY_OLD" -eq 1 ]; then
    PIN="pyinstaller==5.13.2"      # 兼容 Python 3.7
  else
    PIN="pyinstaller"
  fi
  if pip_install "$PIN"; then
    BUILD_DIR="$(mktemp -d)"
    # --collect-all docx：收集 python-docx 的模板数据文件
    #（PyInstaller 5.x 不会自动打包包数据，缺它会报 default-footer.xml 不存在）
    if (cd "$BUILD_DIR" && "$PY" -m PyInstaller --noconfirm --clean \
          --onefile --windowed --collect-all tkinterdnd2 --collect-all docx \
          --specpath "$BUILD_DIR" --distpath "$INSTALL_DIR/bin" \
          --workpath "$BUILD_DIR/build" \
          --name "$APP_NAME" "$INSTALL_DIR/app/app.py") >/dev/null 2>&1; then
      echo "  · 独立程序：$BIN"
    else
      warn "打包未成功（可能缺 binutils 等），自动改用脚本直跑模式（功能完全相同）"
      NO_BUILD=1
    fi
    rm -rf "$BUILD_DIR"
  else
    warn "PyInstaller 未装成，改用脚本直跑模式（功能完全相同）"
    NO_BUILD=1
  fi
fi

if [ "$NO_BUILD" -eq 0 ] && [ -x "$BIN" ]; then
  EXEC="$BIN"
else
  EXEC="$PY $INSTALL_DIR/app/app.py"
fi

# ── 9. 桌面图标 ──
step "注册桌面图标…"
mkdir -p "$HOME/.local/share/applications"
cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Name=$APP_NAME
Name[zh_CN]=$APP_NAME
GenericName=公文排版
GenericName[zh_CN]=公文排版
Comment=按《市委办公室材料排版基本格式》一键排版 Word/WPS 文档
Comment[zh_CN]=按《市委办公室材料排版基本格式》一键排版 Word/WPS 文档
Exec=$EXEC
Icon=$INSTALL_DIR/share/icon.svg
Terminal=false
Categories=Office;WordProcessor;
StartupWMClass=$APP_NAME
EOF
update-desktop-database "$HOME/.local/share/applications" >/dev/null 2>&1 || true

echo
echo "✔ 安装完成！（以用户 $(whoami) 安装）"
echo "  · 启动：应用启动器搜索「材料排版工具」"
echo "  · 终端：$EXEC"
echo "  · 命令行排版：$EXEC --cli 文件.docx"
if [ -f "$SCRIPT_DIR/uninstall.sh" ]; then
  echo "  · 卸载：bash $SCRIPT_DIR/uninstall.sh"
fi
if [ "$(id -u)" -eq 0 ] && [ "${HOME#/root}" != "$HOME" ]; then
  warn "当前以 root 安装；若桌面以其他用户登录导致启动器无图标，请换桌面登录用户重跑本脚本"
fi
