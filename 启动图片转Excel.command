#!/bin/zsh
cd "$(dirname "$0")" || exit 1

if [ ! -x ".venv/bin/python" ]; then
    echo "❌ 未找到运行环境（.venv）。"
    echo "请先用 Python 3.10+ 重建环境："
    echo "   python3 -m venv .venv"
    echo "   source .venv/bin/activate"
    echo "   pip install -r requirements.txt"
    read -r
    exit 1
fi

# 后台启动图形界面，终端窗口随即关闭
nohup .venv/bin/python gui.py >/dev/null 2>&1 &
exit 0
