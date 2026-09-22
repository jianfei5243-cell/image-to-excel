#!/bin/zsh
lsof -ti:8501 | xargs kill 2>/dev/null
rm -f /tmp/image2excel_web.pid
echo "网页版已停止（如浏览器仍开着，可手动关闭标签页）。"
sleep 1
