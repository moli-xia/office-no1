#!/usr/bin/env bash
# 材料排版工具 · 卸载
set -euo pipefail
rm -rf "$HOME/.local/opt/clpb"
rm -f "$HOME/.local/share/applications/clpb.desktop"
update-desktop-database "$HOME/.local/share/applications" 2>/dev/null || true
echo "✔ 已卸载「材料排版工具」（未动系统任何组件与用户文档）"
