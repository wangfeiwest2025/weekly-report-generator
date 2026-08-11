#!/bin/bash
# Linux部署脚本 - 执法处周报材料生成器

set -e

echo "========================================"
echo "开始部署到Linux系统"
echo "========================================"

# 检查Python版本
echo ""
echo "检查Python版本..."
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到Python3，请先安装Python3"
    exit 1
fi

PYTHON_VERSION=$(python3 --version | cut -d' ' -f2)
echo "Python版本：$PYTHON_VERSION"

# 检查pip
echo ""
echo "检查pip..."
if ! command -v pip3 &> /dev/null; then
    echo "错误：未找到pip3，请先安装pip3"
    exit 1
fi

# 创建虚拟环境（推荐）
echo ""
echo "创建Python虚拟环境..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "虚拟环境创建成功"
else
    echo "虚拟环境已存在"
fi

# 激活虚拟环境
echo ""
echo "激活虚拟环境..."
source venv/bin/activate

# 升级pip
echo ""
echo "升级pip..."
pip install --upgrade pip

# 安装依赖
echo ""
echo "安装依赖包..."
pip install -r requirements.txt

# 创建必要的目录
echo ""
echo "创建必要的目录..."
mkdir -p data
mkdir -p logs

# 设置权限
echo ""
echo "设置文件权限..."
chmod +x start.sh
chmod +x stop.sh
chmod 755 streamlit_app.py
chmod 755 weekly_report.py

echo ""
echo "========================================"
echo "部署完成！"
echo "========================================"
echo ""
echo "启动方式："
echo "  ./start.sh"
echo ""
echo "访问地址："
echo "  http://localhost:8501"
echo ""
echo "后台运行："
echo "  nohup ./start.sh > logs/app.log 2>&1 &"
echo ""
echo "========================================"
