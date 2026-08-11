# 执法处周报材料生成器 - 部署说明

## 📦 文件夹结构

```
deploy/
├── web_app/          # 网页应用（推荐）
│   ├── streamlit_app.py     # 网页应用主程序
│   ├── weekly_report.py     # 核心计算逻辑
│   ├── requirements.txt     # 依赖清单
│   ├── start.bat            # Windows启动脚本
│   ├── deploy.sh            # Linux部署脚本
│   ├── start.sh             # Linux启动脚本
│   ├── stop.sh              # Linux停止脚本
│   ├── Dockerfile           # Docker镜像
│   ├── docker-compose.yml   # Docker Compose配置
│   └── weekly-report.service # Systemd服务
├── docs/             # 文档
│   ├── 启动指南.md
│   └── 网页应用使用说明.md
└── README.md         # 本文件
```

## 🚀 快速开始

### 网页应用（推荐）

```bash
cd web_app
start.bat                    # Windows
# 或
chmod +x deploy.sh && ./deploy.sh   # Linux
nohup ./start.sh > logs/app.log 2>&1 &   # Linux后台运行
```

浏览器访问 `http://localhost:8501`

## 📦 部署说明

### Linux部署

#### 方式1：Docker Compose（最简单）

```bash
scp -r web_app/ user@server:/opt/weekly-report/
cd /opt/weekly-report
mkdir -p data
docker-compose up -d
```

#### 方式2：直接部署

```bash
scp -r web_app/ user@server:/opt/weekly-report/
cd /opt/weekly-report
chmod +x deploy.sh
./deploy.sh
nohup ./start.sh > logs/app.log 2>&1 &
```

#### 方式3：Systemd服务

```bash
cd /opt/weekly-report
python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt
sudo cp weekly-report.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now weekly-report
```

详细文档：[web_app/LINUX_DEPLOY.md](web_app/LINUX_DEPLOY.md)

### Windows部署

```bash
cd web_app
start.bat
```

## ⚠️ 注意事项

1. **固定数据已内置**：2025年建筑市场检查.xls、2025年工程建设领域处罚.xls、基本口径.xls 随程序部署，无需额外上传
2. **输入数据**：每周通过网页上传检查原始数据导出和处罚原始数据导出（.xls格式）
3. **历史记录**：写入 `data/周报历史记录.json`（Docker/Systemd方式自动持久化到 data 目录）
4. **默认端口**：8501，可修改
5. **时区**：Docker方式已设置 `TZ=Asia/Shanghai`

---

**版本：** 1.0  
**创建日期：** 2026年8月11日
