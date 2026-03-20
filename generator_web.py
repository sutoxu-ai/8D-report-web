import streamlit as st
from io import BytesIO
from datetime import datetime
import re

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

import openai

# =========================================================
# 基础配置
# =========================================================
st.set_page_config(
    page_title="8D报告自动生成系统",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# 从 Streamlit Secrets 读取 API
# =========================================================
API_KEY = st.secrets["DEEPSEEK_API_KEY"]
BASE_URL = st.secrets["DEEPSEEK_BASE_URL"]

client = openai.OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL
)

# =========================================================
# Web 授权状态（Session 级）
# =========================================================
if "license_type" not in st.session_state:
    st.session_state.license_type = "trial"   # trial / active

if "history" not in st.session_state:
    st.session_state.history = []

# =========================================================
# 授权区
# =========================================================
def license_panel():
    with st.sidebar:
        st.header("系统状态")

        if st.session_state.license_type == "trial":
            st.warning("⚠️ 当前为试用版（仅预览，不可导出 Word）")
            code = st.text_input("输入激活码", type="password")
            if st.button("激活正式版"):
                if len(code) >= 8:
                    st.session_state.license_type = "active"
                    st.success("✅ 激活成功（当前会话）")
                else:
                    st.error("激活码格式无效")
        else:
            st.success("✅ 正式版（当前会话）")

        st.markdown("---")
        st.header("历史记录（最近 5 条）")

        for i, item in enumerate(st.session_state.history[-5:]):
            if st.button(f"📄 {item['title']}", key=f"h_{i}"):
                st.session_state.current_result = item["content"]

# =========================================================
# Word 写入
# =========================================================
def write_8d_to_word(doc, raw_text):
    lines = raw_text.split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            continue

        line = re.sub(r"^[\-\*\•\d\.\)\s]+", "", line)
        line = line.replace("**", "")

        if re.match(r"^D[1-8]\b", line):
            p = doc.add_heading(line, level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p = doc.add_paragraph(line)
            p.paragraph_format.space_after = Pt(6)

# =========================================================
# 主界面
# =========================================================
st.title("📊 8D 报告自动生成系统")
st.markdown("---")

license_panel()

col1, col2 = st.columns([1, 1])

# ---------------- 输入区 ----------------
with col1:
    st.header("📝 输入基本信息")

    p_desc = st.text_area(
        "不良信息描述",
        height=160,
        placeholder="请使用 5W2H 描述异常事实"
    )

    p_name = st.text_input("产品型号 / 名称")
    cust = st.text_input("客户名称")

    r1, r2 = st.columns(2)
    o_date = r1.date_input("发现日期", datetime.now())
    qty = r2.number_input("不良数量", min_value=1)

    if st.button("🚀 自动生成 8D 报告", use_container_width=True):
        if not p_desc:
            st.error("请先输入异常描述")
        else:
            with st.spinner("8D 报告生成中（约 60 秒）"):
                sys_prompt = """
你将直接输出一份【正式 8D 报告正文】。
【严格要求】：
- 禁止任何自我介绍或第一人称
- 所有内容必须可直接提交客户
- 不允许使用问号

必须完整包含：
D1. 成立团队
D2. 问题描述（5W2H）
D3. 临时围堵措施
D4. 根本原因分析
  A. Why Occurrence（4M1E + 仅异常项 5Why）
  B. Why Escape（5Why）
D5. 永久纠正措施选择
D6. 实施纠正措施
D7. 防止再发生（标准化 / FMEA / 控制计划）
D8. 团队表彰

A 与 B 之间空两行。
"""

                resp = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": sys_prompt},
                        {"role": "user", "content": f"异常事实：{p_desc}"}
                    ],
                    temperature=0.2
                )

                result = resp.choices[0].message.content
                st.session_state.current_result = result
                st.session_state.history.append({
                    "title": f"{p_name[:10]}_{o_date}",
                    "content": result
                })

# ---------------- 预览 / 导出区 ----------------
with col2:
    st.header("📄 8D 报告预览")

    if "current_result" in st.session_state:
        content = st.session_state.current_result
        st.markdown(content)

        if st.session_state.license_type == "active":
            doc = Document()
            doc.styles["Normal"].font.name = "宋体"
            doc.styles["Normal"]._element.rPr.rFonts.set(
                qn("w:eastAsia"), "宋体"
            )

            doc.add_heading("8D 问题纠正与预防措施报告", 0)
            write_8d_to_word(doc, content)

            bio = BytesIO()
            doc.save(bio)

            st.download_button(
                "📥 导出 Word 报告",
                bio.getvalue(),
                file_name=f"8D_Report_{p_name}.docx",
                use_container_width=True
            )
        else:
            st.info("🔒 试用版仅支持预览，激活后可导出 Word")