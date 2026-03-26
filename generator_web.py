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
# 0. Page config
# =========================================================
st.set_page_config(
    page_title="8D Report Generator",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# 1. Language state
# =========================================================
if "lang" not in st.session_state:
    st.session_state.lang = "zh"  # zh / en

TEXT = {
    "zh": {
        "title": "📊 8D 报告自动生成系统",
        "status": "系统状态",
        "trial": "⚠️ 试用版",
        "active": "✅ 正式版",
        "activate": "🔑 授权 / 续费",
        "input": "📝 输入基本信息",
        "desc": "不良信息描述",
        "generate": "🚀 自动生成 8D 报告",
        "preview": "📄 8D 报告预览",
        "export": "📥 导出 Word 报告",
        "expired": "授权已到期",
        "trial_left": "剩余试用次数",
        "word_title": "8D 问题纠正与预防措施报告"
    },
    "en": {
        "title": "📊 8D Report Auto Generator",
        "status": "System Status",
        "trial": "⚠️ Trial Version",
        "active": "✅ Licensed Version",
        "activate": "🔑 License / Renewal",
        "input": "📝 Input Information",
        "desc": "Problem Description",
        "generate": "🚀 Generate 8D Report",
        "preview": "📄 8D Report Preview",
        "export": "📥 Export Word Report",
        "expired": "License expired",
        "trial_left": "Remaining trial runs",
        "word_title": "8D Corrective and Preventive Action Report"
    }
}

def T(key: str) -> str:
    return TEXT[st.session_state.lang].get(key, key)

# =========================================================
# 2. API (DeepSeek / OpenAI compatible)
# =========================================================
API_KEY = st.secrets["DEEPSEEK_API_KEY"]
BASE_URL = st.secrets["DEEPSEEK_BASE_URL"]

client = openai.OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL
)

# =========================================================
# 3. License (local file)
# =========================================================
LICENSE_FILE = Path("license.json")
MAX_TRIAL = 3
LICENSE_DAYS = 365

def load_license():
    if LICENSE_FILE.exists():
        try:
            return json.loads(LICENSE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"trial_used": 0, "license_expire": None}

def save_license(data):
    LICENSE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

license_data = load_license()

def trial_remaining():
    return max(0, MAX_TRIAL - license_data.get("trial_used", 0))

def is_license_valid():
    exp = license_data.get("license_expire")
    if not exp:
        return False
    return datetime.today().date() <= datetime.strptime(exp, "%Y-%m-%d").date()

def activate_license():
    exp = datetime.today().date() + timedelta(days=LICENSE_DAYS)
    license_data["license_expire"] = exp.strftime("%Y-%m-%d")
    save_license(license_data)

def inc_trial():
    license_data["trial_used"] += 1
    save_license(license_data)

# =========================================================
# 4. Session
# =========================================================
if "license_type" not in st.session_state:
    st.session_state.license_type = "active" if is_license_valid() else "trial"

if "history" not in st.session_state:
    st.session_state.history = []

def sync_license():
    st.session_state.license_type = "active" if is_license_valid() else "trial"

sync_license()

# =========================================================
# 5. Sidebar
# =========================================================
with st.sidebar:
    st.markdown("### 🌐 Language")
    lang_ui = st.radio(
        "",
        ["中文", "English"],
        index=0 if st.session_state.lang == "zh" else 1,
        horizontal=True
    )
    st.session_state.lang = "zh" if lang_ui == "中文" else "en"

    st.markdown("---")
    st.header(T("status"))

    if st.session_state.license_type == "active":
        st.success(T("active"))
        if license_data.get("license_expire"):
            st.info(f"📅 {license_data['license_expire']}")
    else:
        st.warning(T("trial"))
        st.caption(f"{T('trial_left')}: {trial_remaining()}")

    st.markdown("---")
    st.subheader(T("activate"))

    code = st.text_input("Activation Code", type="password")
    if st.button("Activate / Renew", use_container_width=True):
        if len(code) >= 8:
            activate_license()
            sync_license()
            st.success("Activated")
            st.rerun()
        else:
            st.error("Invalid code")

# =========================================================
# 6. Word writer
# =========================================================
def write_to_word(doc, text):
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        line = re.sub(r"^[\-\*\d\.\)\s]+", "", line)
        if re.match(r"^D[1-8]\b", line):
            doc.add_heading(line, level=1)
        else:
            p = doc.add_paragraph(line)
            p.paragraph_format.space_after = Pt(6)

# =========================================================
# 7. UI
# =========================================================
st.title(T("title"))
st.markdown("---")

col1, col2 = st.columns([1, 1])

with col1:
    st.header(T("input"))
    desc = st.text_area(T("desc"), height=160)

    if st.button(T("generate"), use_container_width=True):
        sync_license()
        if st.session_state.license_type == "trial" and trial_remaining() <= 0:
            st.error("Trial exhausted")
            st.stop()

        if st.session_state.lang == "zh":
            sys_prompt = """
你将直接输出一份【正式 8D 报告正文】。
要求：
- 不使用第一人称
- 内容可直接提交客户
- 不允许问句
必须包含 D1–D8 全结构
"""
        else:
            sys_prompt = """
You will generate a professional 8D report.
Requirements:
- No first-person language
- Customer-ready content
- No questions
Must include full D1–D8 structure
"""

        with st.spinner("Generating..."):
            resp = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": desc}
                ],
                temperature=0.2
            )

            result = resp.choices[0].message.content
            st.session_state.current_result = result
            st.session_state.history.append({
                "title": f"8D_{datetime.now().date()}",
                "content": result
            })

            if st.session_state.license_type == "trial":
                inc_trial()

with col2:
    st.header(T("preview"))
    if "current_result" in st.session_state:
        content = st.session_state.current_result
        st.markdown(content)

        sync_license()
        if st.session_state.license_type == "active":
            doc = Document()
            if st.session_state.lang == "zh":
                doc.styles["Normal"].font.name = "宋体"
                doc.styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            else:
                doc.styles["Normal"].font.name = "Calibri"

            doc.add_heading(T("word_title"), 0)
            write_to_word(doc, content)

            bio = BytesIO()
            doc.save(bio)

            st.download_button(
                T("export"),
                bio.getvalue(),
                file_name="8D_Report.docx",
                use_container_width=True
            )
