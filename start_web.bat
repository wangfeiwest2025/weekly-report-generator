@echo off
chcp 65001 >nul
echo ========================================
echo 启动"执法处周报材料生成器"网页版
echo ========================================
echo.

echo 检查依赖...
cd /d "%~dp0"
pip install streamlit pandas openpyxl xlrd python-docx -q

echo.
echo 启动网页应用...
echo 浏览器自动打开后请选择"检查原始数据导出"和"处罚原始数据导出"文件
echo.
streamlit run "%~dp0streamlit_app.py" --server.port 8501
pause