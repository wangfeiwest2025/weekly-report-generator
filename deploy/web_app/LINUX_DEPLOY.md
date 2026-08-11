# Linux完整部署指南 - 执法处周报材料生成器

## 📦 部署包内容

```
web_app/
├── streamlit_app.py              # 网页应用主程序
├── weekly_report.py              # 核心计算逻辑
├── web_report.py                 # 备用Flask版（可选）
├── requirements.txt              # 依赖清单
├── start.bat                     # Windows启动脚本
├── deploy.sh                     # Linux部署脚本
├── start.sh                      # Linux启动脚本
├── stop.sh                       # Linux停止脚本
├── Dockerfile                    # Docker镜像
├── docker-compose.yml            # Docker Compose配置
├── weekly-report.service         # Systemd服务
├── LINUX_DEPLOY.md               # 本文档
├── QUICKSTART.md                 # Linux快速部署
├── 2025年建筑市场检查.xls         # 内置固定数据（去年同期检查）
├── 2025年工程建设领域处罚.xls      # 内置固定数据（去年同期处罚）
└── 基本口径.xls                   # 内置固定数据（在施工程量/机构映射）
```

## 🚀 环境要求

- Linux (Ubuntu/Debian/CentOS)
- Python 3.9+（建议 3.11）
- pip3 / python3-venv
- Docker（可选，Docker Compose 方式需要）

## 📋 三种部署方式

### 方式1：Docker Compose（推荐，最简单）

```bash
# 1. 上传整个 web_app 文件夹
scp -r web_app/ user@server:/opt/weekly-report/

# 2. 进入目录
cd /opt/weekly-report

# 3. 创建数据目录
mkdir -p data

# 4. 启动（自动构建镜像）
docker-compose up -d

# 5. 访问
# http://your-server:8501
```

### 方式2：直接运行（无需Docker）

```bash
# 1. 上传整个 web_app 文件夹
scp -r web_app/ user@server:/opt/weekly-report/

# 2. 运行部署脚本（创建venv、安装依赖、创建目录）
cd /opt/weekly-report
chmod +x deploy.sh
./deploy.sh

# 3. 后台启动
nohup ./start.sh > logs/app.log 2>&1 &

# 4. 访问
# http://your-server:8501

# 停止
./stop.sh
```

### 方式3：Systemd服务（生产环境推荐）

```bash
# 1. 部署应用（创建venv并安装依赖）
cd /opt/weekly-report
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
mkdir -p data logs

# 2. 配置服务（注意用户权限，需要修改service中的User）
sudo mkdir -p /opt/weekly-report/data
sudo chown -R www-data:www-data /opt/weekly-report
sudo cp weekly-report.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable weekly-report
sudo systemctl start weekly-report

# 3. 查看状态
sudo systemctl status weekly-report

# 4. 查看日志
sudo journalctl -u weekly-report -f
```

## 🔧 常用命令

### Docker方式
```bash
docker-compose up -d          # 启动
docker-compose stop           # 停止
docker-compose restart        # 重启
docker-compose logs -f        # 查看日志
docker-compose up -d --build  # 更新代码后重建
```

### 直接运行
```bash
nohup ./start.sh > logs/app.log 2>&1 &   # 启动
./stop.sh                                 # 停止
tail -f logs/app.log                      # 查看日志
```

### Systemd
```bash
sudo systemctl start|stop|restart weekly-report
sudo systemctl status weekly-report
sudo journalctl -u weekly-report -f
```

## 🌐 Nginx反向代理（可选）

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
    }
}
```

```bash
sudo ln -s /etc/nginx/sites-available/weekly-report /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

## 🔥 防火墙配置

```bash
# Ubuntu/Debian (UFW)
sudo ufw allow 8501/tcp
sudo ufw reload

# CentOS (Firewalld)
sudo firewall-cmd --permanent --add-port=8501/tcp
sudo firewall-cmd --reload
```

## 💾 数据持久化

- **历史记录**（上周数值，用于环比一致性）：写入 `data/周报历史记录.json`
  - 直接运行/Systemd：`start.sh` 自动设置 `REPORT_DATA_DIR` 指向 `./data`
  - Docker：`docker-compose.yml` 挂载 `./data:/app/data`
- **固定数据文件**：随程序内置，无需额外准备
- **上传的原始导出文件**：仅存内存/临时目录，生成后自动清理，不会落盘

## ❓ 常见问题

### 端口被占用
```bash
sudo lsof -i:8501
sudo kill -9 <PID>
```

### 权限问题
```bash
chmod +x *.sh
chmod -R 755 /opt/weekly-report
```

### Docker构建失败
```bash
# 查看日志
docker logs weekly-report-app

# 重新构建（不使用缓存）
docker-compose build --no-cache
docker-compose up -d
```

### 中文文件名乱码
Linux 默认 UTF-8 编码，正常情况下无问题。若使用非 UTF-8 的 SSH/SCP 客户端上传文件，请检查文件名编码。

---

**版本：** 1.0  
**创建日期：** 2026年8月11日
