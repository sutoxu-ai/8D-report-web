import streamlit as st
from io import BytesIO
from datetime import datetime, timedelta
import re
import json
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import openai

# =========================================================
# 0. 页面配置
# =========================================================
st.set_page_config(
    page_title="8D报告自动生成系统",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# 1. Secrets：DeepSeek
# =========================================================
API_KEY = st.secrets["DEEPSEEK_API_KEY"]
BASE_URL = st.secrets["DEEPSEEK_BASE_URL"]

client = openai.OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL
)

# =========================================================
# 2. License 本地持久化（替代 Cookie）
# =========================================================
LICENSE_FILE = Path("license.json")
MAX_TRIAL_GENERATIONS = 3
LICENSE_DAYS = 365

def load_license():
    if LICENSE_FILE.exists():
        try:
            return json.loads(LICENSE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "trial_used": 0,
        "license_expire": None
    }

def save_license(data: dict):
    LICENSE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

license_data = load_license()

def trial_remaining():
    return max(0, MAX_TRIAL_GENERATIONS - license_data.get("trial_used", 0))

def is_license_valid_today():
    exp = license_data.get("license_expire")
    if not exp:
        return False
    return datetime.today().date() <= datetime.strptime(exp, "%Y-%m-%d").date()

def activate_new_license():
    exp_date = datetime.today().date() + timedelta(days=LICENSE_DAYS)
    license_data["license_expire"] = exp_date.strftime("%Y-%m-%d")
    save_license(license_data)

def inc_trial_used():
    license_data["trial_used"] = license_data.get("trial_used", 0) + 1
    save_license(license_data)

# =========================================================
# 3. Session State
# =========================================================
if "license_type" not in st.session_state:
    st.session_state.license_type = "active" if is_license_valid_today() else "trial"

if "history" not in st.session_state:
    st.session_state.history = []

def sync_license_state():
    st.session_state.license_type = "active" if is_license_valid_today() else "trial"

sync_license_state()

# =========================================================
# 4. Sidebar：授权面板
# =========================================================
def license_panel():
    with st.sidebar:
        st.header("系统状态")

        exp = license_data.get("license_expire")

        if st.session_state.license_type == "active":
            st.success("✅ 正式版：无限次生成 + 支持 Word 导出")
            if exp:
                st.info(f"📅 授权有效期至：{exp}")
        else:
            used = license_data.get("trial_used", 0)
            rem = trial_remaining()
            st.warning(f"⚠️ 试用版：允许生成 {MAX_TRIAL_GENERATIONS} 次（已用 {used}，剩余 {rem}）")

            if exp:
                st.error(f"❌ 授权已到期（到期日：{exp}）")

            if rem <= 0:
                st.error("❌ 试用次数已用完，请激活。")

        st.markdown("---")
        st.subheader("🔑 授权 / 续费")

        code = st.text_input("输入激活码", type="password")

        if st.button("立即激活 / 续费", use_container_width=True):
            if len(code) >= 8:
                activate_new_license()
                sync_license_state()
                st.success("✅ 授权成功，已生效一年")
                st.rerun()
            else:
                st.error("激活码无效（至少 8 位）")

        st.markdown("---")
        st.header("历史记录（最近 5 条）")

        for i, item in enumerate(st.session_state.history[-5:]):
            if st.button(f"📄 {item['title']}", key=f"h_{i}", use_container_width=True):
                st.session_state.current_result = item["content"]

# =========================================================
# 5. Word 写入
# =========================================================
def write_8d_to_word(doc, raw_text):
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        line = re.sub(r"^[\-\*\•\d\.\)\s]+", "", line)
        line = line.replace("**", "")
        if re.match(r"^D[1-8]\b", line):
            doc.add_heading(line, level=1)
        else:
            p = doc.add_paragraph(line)
            p.paragraph_format.space_after = Pt(6)

# =========================================================
# 6. UI
# =========================================================
st.title("📊 8D 报告自动生成系统")
st.markdown("---")
license_panel()

col1, col2 = st.columns([1, 1])

with col1:
    st.header("📝 输入基本信息")

    p_desc = st.text_area("不良信息描述", height=160)
    p_name = st.text_input("产品型号 / 名称")
    cust = st.text_input("客户名称")

    r1, r2 = st.columns(2)
    o_date = r1.date_input("发现日期", datetime.now())
    qty = r2.number_input("不良数量", min_value=1)

    if st.button("🚀 自动生成 8D 报告", use_container_width=True):
        sync_license_state()

        if st.session_state.license_type == "trial" and trial_remaining() <= 0:
            st.error("❌ 试用次数已用完")
            st.stop()

        with st.spinner("8D 报告生成中…"):
            resp = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你将直接输出正式 8D 报告正文"},
                    {"role": "user", "content": p_desc}
                ],
                temperature=0.2
            )

            result = resp.choices[0].message.content
            st.session_state.current_result = result

            st.session_state.history.append({
                "title": f"{p_name or '8D'}_{o_date}",
                "content": result
            })

            if st.session_state.license_type == "trial":
                inc_trial_used()
                st.info(f"ℹ️ 剩余试用次数：{trial_remaining()}")

with col2:
    st.header("📄 8D 报告预览")

    if "current_result" in st.session_state:
        content = st.session_state.current_result
        st.markdown(content)

        sync_license_state()

        if st.session_state.license_type == "active":
            doc = Document()
            doc.styles["Normal"].font.name = "宋体"
            doc.styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            doc.add_heading("8D 问题纠正与预防措施报告", 0)
            write_8d_to_word(doc, content)

            bio = BytesIO()
            doc.save(bio)

            st.download_button(
                "📥 导出 Word 报告",
                bio.getvalue(),
                file_name=f"8D_Report_{p_name or 'NA'}.docx",
                use_container_width=True
            )
        else:
            st.info("🔒 激活后可导出 Word")

