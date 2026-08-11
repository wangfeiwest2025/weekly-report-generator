#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
"执法处周报材料生成器" —— Streamlit 网页版
===========================================

参考 deploy 目录的界面风格与部署方式：
- 上传检查原始数据导出 + 处罚原始数据导出（不再需要选择固定数据文件）
- 侧边栏设置参数
- 自动填充时间（根据导出文件数据）
- 进度指示
- 预览正文并下载 docx

与 weekly_report.py 共享计算逻辑。

运行：双击 start_web.bat 或 streamlit run streamlit_app.py
"""

import os
import sys
import tempfile
import datetime as dt
import traceback

import streamlit as st
import weekly_report as wr

# ---------------------------------------------------------------------------
# 页面配置
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="执法处周报材料生成器",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="expanded",
)

# 自定义 CSS
st.markdown(
    """
<style>
    .main-header {
        font-size: 32px; font-weight: bold; color: #1f77b4;
        text-align: center; margin: 10px 0 20px;
    }
    .section-header {
        font-size: 20px; font-weight: bold; color: #333;
        margin-top: 20px; margin-bottom: 10px;
    }
    .para-box {
        background: #fbfbf6; border: 1px solid #e5e1cf; border-radius: 6px;
        padding: 12px 14px; font-size: 15px; line-height: 1.9;
        margin-bottom: 10px; text-indent: 2em;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 标题
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="main-header">📋 执法处周报材料生成器 <span style="font-size:15px;color:#999">Streamlit版</span></div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# 侧边栏
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 设置")
    st.markdown("---")

    st.subheader("📅 时间范围")
    auto_d0 = st.session_state.get("_auto_d0")
    auto_d1 = st.session_state.get("_auto_d1")
    d0 = st.date_input("开始日期（本周周一）", value=auto_d0)
    d1 = st.date_input("结束日期（本周周日）", value=auto_d1)

    st.subheader("📝 上周处罚件数")
    lw_punish = st.number_input(
        "留空则自动计算", value=None, min_value=0, step=1, format="%d"
    )

    st.markdown("---")
    st.subheader("📖 使用说明")
    st.info(
        """
    1. 上传检查原始数据导出文件
    2. 上传处罚原始数据导出文件
    3. 设置时间范围（上传文件后自动填充）
    4. 点击"生成周报"
    5. 预览正文并下载 docx
    """
    )
    st.caption("固定数据文件已内置于程序文件夹")
    st.markdown("---")
    st.markdown(
        '<div style="text-align:center;color:#999;font-size:12px">© 2026 执法处周报材料生成器</div>',
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# 文件上传区
# ---------------------------------------------------------------------------
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.markdown("**📤 检查原始数据导出**")
    check_file = st.file_uploader(
        "选择检查数据（.xls）", type=["xls", "xlsx"], label_visibility="collapsed", key="check"
    )
with col2:
    st.markdown("**📤 处罚原始数据导出**")
    punish_file = st.file_uploader(
        "选择处罚数据（.xls）", type=["xls", "xlsx"], label_visibility="collapsed", key="punish"
    )

# ---------------------------------------------------------------------------
# 自动填充时间（文件初次上传时从导出数据推断）
# ---------------------------------------------------------------------------
upload_key = (
    (check_file.name, punish_file.name)
    if check_file and punish_file
    else None
)
if upload_key and st.session_state.get("_last_upload_key") != upload_key:
    st.session_state["_last_upload_key"] = upload_key
    try:
        with tempfile.TemporaryDirectory() as tmp:
            p1 = os.path.join(tmp, check_file.name)
            p2 = os.path.join(tmp, punish_file.name)
            with open(p1, "wb") as f:
                f.write(check_file.getbuffer())
            with open(p2, "wb") as f:
                f.write(punish_file.getbuffer())
            wk = wr.week_from_exports(p1, p2)
            if wk:
                st.session_state["_auto_d0"] = wk[0]
                st.session_state["_auto_d1"] = wk[1]
                st.rerun()
    except Exception:
        pass

# ---------------------------------------------------------------------------
# 上传状态提示
# ---------------------------------------------------------------------------
if not check_file or not punish_file:
    st.warning("⚠️ 请先上传检查原始数据导出和处罚原始数据导出文件")
else:
    col_a, col_b = st.columns(2)
    with col_a:
        st.success(f"✓ 已上传：{check_file.name}")
    with col_b:
        st.success(f"✓ 已上传：{punish_file.name}")

# 提示未设置日期
if check_file and punish_file and (d0 is None or d1 is None):
    st.info("📅 请在左侧侧边栏设置时间范围（或等待文件读取后自动填充）")

# ---------------------------------------------------------------------------
# 生成周报
# ---------------------------------------------------------------------------
if check_file and punish_file and d0 and d1:
    st.divider()

    if st.button("🚀 生成周报", type="primary"):
        logs = []
        docx_bytes = None
        para1 = para2 = ""
        out_name = f"执法处周报材料({d0.month}月{d0.day}日-{d1.month}月{d1.day}日).docx"

        try:
            with st.status("⏳ 正在处理数据，请稍候...", expanded=True) as status:
                st.write("保存上传文件...")
                # 用临时目录存放上传文件
                tmpdir = tempfile.mkdtemp(prefix="weekly_report_")
                insp_path = os.path.join(tmpdir, check_file.name)
                pen_path = os.path.join(tmpdir, punish_file.name)
                out_path = os.path.join(tmpdir, out_name)
                with open(insp_path, "wb") as f:
                    f.write(check_file.getbuffer())
                with open(pen_path, "wb") as f:
                    f.write(punish_file.getbuffer())

                st.write("读取固定数据文件...")
                # run_report 内部会读取 2025年建筑市场检查.xls / 工程建设领域处罚.xls / 基本口径.xls

                st.write("计算本周/上周/去年同期统计...")
                params = {
                    "check_raw": insp_path,
                    "punish_raw": pen_path,
                    "d0": d0.isoformat(),
                    "d1": d1.isoformat(),
                    "lw_punish": str(int(lw_punish)) if lw_punish else "",
                    "out": out_path,
                }
                para1, para2, out_path = wr.run_report(params, logs.append)

                st.write("生成 docx 文件...")
                with open(out_path, "rb") as f:
                    docx_bytes = f.read()

                # 清理临时目录
                import shutil

                shutil.rmtree(tmpdir, ignore_errors=True)

                status.update(label="✅ 周报生成成功！", state="complete")

        except Exception as e:
            st.error(f"❌ 生成失败：{e}")
            with st.expander("📋 详细错误"):
                st.code("\n".join(logs))
                st.code(traceback.format_exc())

        # 显示结果
        if docx_bytes:
            st.divider()
            st.success("✅ 周报生成成功！")

            st.markdown("#### 📄 正文一")
            st.markdown(f'<div class="para-box">{para1}</div>', unsafe_allow_html=True)

            st.markdown("#### 📄 正文二")
            st.markdown(f'<div class="para-box">{para2}</div>', unsafe_allow_html=True)

            st.download_button(
                label="📥 下载 docx 文件",
                data=docx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document",
            )

            with st.expander("📋 运行日志"):
                st.text("\n".join(logs))