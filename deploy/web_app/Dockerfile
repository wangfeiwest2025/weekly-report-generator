# 使用Python 3.11作为基础镜像
FROM python:3.11-slim

# 设置工作目录
WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用文件（含内置固定数据）
COPY streamlit_app.py weekly_report.py ./
COPY 2025年建筑市场检查.xls 2025年工程建设领域处罚.xls 基本口径.xls ./

# 创建必要的目录
RUN mkdir -p /app/data

# 数据目录（历史记录持久化，用 docker-compose 挂载 volume 覆盖）
ENV REPORT_DATA_DIR=/app/data

# 暴露端口
EXPOSE 8501

# 健康检查
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:8501/_stcore/health || exit 1

# 启动应用
CMD ["streamlit", "run", "streamlit_app.py", "--server.address", "0.0.0.0", "--server.port", "8501", "--browser.gatherUsageStats", "false"]
