# 执法处周报材料生成器 - 部署包

## 🎉 欢迎使用

这是一个组织良好的部署包，包含网页应用，方便您快速部署到任何环境（Windows / Linux / Docker）。

## 📁 文件夹结构

```
deploy/
├── README.md                 # 本文件
├── DEPLOY.md                 # 部署说明
├── 结构说明.md               # 文件夹结构详细说明
│
├── web_app/                  # ⭐ 网页应用（推荐）
│   ├── streamlit_app.py     # 网页应用主程序（Streamlit）
│   ├── weekly_report.py     # 核心计算逻辑（与桌面版共用）
│   ├── web_report.py        # 备用Flask版（可选）
│   ├── requirements.txt     # 依赖清单
│   ├── start.bat            # Windows启动脚本
│   ├── deploy.sh            # Linux部署脚本
│   ├── start.sh             # Linux启动脚本
│   ├── stop.sh              # Linux停止脚本
│   ├── Dockerfile           # Docker镜像
│   ├── docker-compose.yml   # Docker Compose配置
│   ├── weekly-report.service # Systemd服务
│   ├── LINUX_DEPLOY.md      # Linux部署详细指南
│   ├── QUICKSTART.md        # Linux快速部署
│   ├── 2025年建筑市场检查.xls       # 内置固定数据（去年同期检查）
│   ├── 2025年工程建设领域处罚.xls    # 内置固定数据（去年同期处罚）
│   └── 基本口径.xls                 # 内置固定数据（在施工程量/机构映射）
│
└── docs/                     # 📖 文档目录
    ├── 启动指南.md
    └── 网页应用使用说明.md
```

## 🚀 快速开始

### Windows

```bash
cd web_app
start.bat
```

### Linux（三种方式任选）

```bash
# Docker Compose（推荐）
cd web_app
mkdir -p data
docker-compose up -d

# 直接运行
cd web_app
chmod +x deploy.sh && ./deploy.sh
nohup ./start.sh > logs/app.log 2>&1 &

# Systemd服务
cd web_app
sudo cp weekly-report.service /etc/systemd/system/
sudo systemctl enable --now weekly-report
```

浏览器访问：`http://localhost:8501`

## 📖 使用说明

1. 上传"检查原始数据导出"文件（如 检查数据0803.xls）
2. 上传"处罚原始数据导出"文件（如 处罚数据0803.xls）
3. 时间范围自动填充（也可手动调整）
4. 点击"🚀 生成周报"
5. 预览正文并下载 docx 文件

**固定数据文件已内置**（2025年同期检查/处罚数据、基本口径），无需上传。

## 📦 部署到生产环境

### Linux部署（推荐）

详细文档：[web_app/LINUX_DEPLOY.md](web_app/LINUX_DEPLOY.md) 和 [web_app/QUICKSTART.md](web_app/QUICKSTART.md)

### Windows部署

```bash
cd web_app
start.bat
```

## 🔧 环境要求

| 方式 | 环境要求 |
|------|---------|
| 网页应用 | Python 3.9+, Streamlit, Pandas, OpenPyXL, xlrd, python-docx, lxml |
| Docker | Docker 20+, Docker Compose |

## ⚠️ 常见问题

### Q1: 无法启动应用
- 检查是否安装了所有依赖包（`pip install -r requirements.txt`）
- 检查端口是否被占用（默认8501）
- 查看错误日志

### Q2: 生成的文件无法打开
- 使用 Word 2007 或更高版本打开 docx 文件
- 检查下载是否完整

### Q3: Linux下中文文件名乱码
- 确保 SSH/SCP 客户端使用 UTF-8 编码上传文件

## 📞 技术支持

详细说明请查看 `docs/` 目录下的文档。

---

**版本：** 1.0  
**创建日期：** 2026年8月11日  
**作者：** Qoder
