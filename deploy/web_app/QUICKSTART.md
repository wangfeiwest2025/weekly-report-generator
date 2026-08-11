# Linux快速部署指南 - 执法处周报材料生成器

## 🚀 30秒快速部署

### 方式1：Docker Compose（推荐）

```bash
# 1. 上传文件
scp -r web_app/ user@server:/opt/weekly-report/

# 2. 进入目录
cd /opt/weekly-report

# 3. 创建数据目录
mkdir -p data

# 4. 启动
docker-compose up -d

# 5. 访问
# http://your-server:8501
```

### 方式2：直接运行

```bash
# 1. 上传文件
scp -r web_app/ user@server:/opt/weekly-report/

# 2. 运行部署脚本
cd /opt/weekly-report
chmod +x deploy.sh
./deploy.sh

# 3. 后台启动
nohup ./start.sh > logs/app.log 2>&1 &

# 4. 访问
# http://your-server:8501
```

### 方式3：Systemd服务（生产环境）

```bash
# 1. 部署应用
cd /opt/weekly-report
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
mkdir -p data logs

# 2. 配置服务
sudo chown -R www-data:www-data /opt/weekly-report
sudo cp weekly-report.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable weekly-report
sudo systemctl start weekly-report

# 3. 查看状态
sudo systemctl status weekly-report
```

---

## 🔧 常用命令

| 操作 | Docker | 直接运行 | Systemd |
|------|--------|----------|---------|
| 启动 | `docker-compose up -d` | `nohup ./start.sh > logs/app.log 2>&1 &` | `sudo systemctl start weekly-report` |
| 停止 | `docker-compose stop` | `./stop.sh` | `sudo systemctl stop weekly-report` |
| 重启 | `docker-compose restart` | 停止+启动 | `sudo systemctl restart weekly-report` |
| 日志 | `docker-compose logs -f` | `tail -f logs/app.log` | `sudo journalctl -u weekly-report -f` |

## 📊 使用方式

部署完成后，浏览器打开 `http://your-server:8501`：

1. 上传"检查原始数据导出"（如 检查数据0803.xls）
2. 上传"处罚原始数据导出"（如 处罚数据0803.xls）
3. 时间范围自动填充（也可手动调整）
4. 点击"🚀 生成周报"
5. 预览正文并下载 docx

固定数据文件（2025年同期数据、基本口径）已内置，无需上传。

## ✅ 验证部署

```bash
# 检查进程
ps aux | grep streamlit

# 检查端口
netstat -tlnp | grep 8501

# 测试访问
curl http://localhost:8501
```

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

### Docker问题
```bash
docker logs weekly-report-app
docker-compose up -d --build
```

---

**快速开始：** 上传文件 → 运行deploy.sh（或docker-compose up -d）→ 访问8501端口
