#!/bin/zsh
cd "$(dirname "$0")" || exit 1

PORT=8501
URL="http://localhost:$PORT"

if [ ! -x ".venv/bin/python" ]; then
    echo "❌ 未找到运行环境（.venv）。"
    echo "请先安装 Python 3.10+，然后按 README「本地运行网页版」创建环境。"
    read -r
    exit 1
fi

# 已经在运行：直接打开浏览器
if curl -s -o /dev/null --max-time 2 "$URL"; then
    open "$URL"
    exit 0
fi

echo "正在启动网页版，稍等，浏览器将自动打开…"

# 后台启动服务
nohup .venv/bin/python -m streamlit run app.py --server.headless true --server.port "$PORT" >/tmp/image2excel_web.log 2>&1 &
echo $! > /tmp/image2excel_web.pid

# 等待服务就绪后打开浏览器
for _ in {1..30}; do
    if curl -s -o /dev/null --max-time 2 "$URL"; then
        open "$URL"
        exit 0
    fi
    sleep 1
done

echo "❌ 启动超时，请查看日志 /tmp/image2excel_web.log"
read -r
exit 1
