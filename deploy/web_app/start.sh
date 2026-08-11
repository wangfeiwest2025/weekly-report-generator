#!/bin/bash
# 启动脚本 - 执法处周报材料生成器

# 激活虚拟环境
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# 数据目录（历史记录持久化，可选）
if [ -d "data" ]; then
    export REPORT_DATA_DIR="$(pwd)/data"
fi

# 启动Streamlit应用
echo "启动网页应用..."
streamlit run streamlit_app.py \
    --server.address 0.0.0.0 \
    --server.port 8501 \
    --browser.gatherUsageStats false
