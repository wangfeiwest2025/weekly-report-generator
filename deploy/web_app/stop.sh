#!/bin/bash
# 停止脚本 - 执法处周报材料生成器

echo "停止网页应用..."

# 查找并停止Streamlit进程
pkill -f "streamlit run streamlit_app.py" || true

# 等待进程结束
sleep 2

# 确认进程已停止
if pgrep -f "streamlit run streamlit_app.py" > /dev/null; then
    echo "强制停止进程..."
    pkill -9 -f "streamlit run streamlit_app.py"
fi

echo "应用已停止"
