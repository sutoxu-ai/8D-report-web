import streamlit as st
from io import BytesIO
from datetime import datetime, timedelta
import re
import json
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import openai

# =========================================================
# 0. 页面配置
# =========================================================
st.set_page_config(
    page_title="8D报告智能生成助手",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# 1. 多语言文本配置
# =========================================================
TEXT = {
    "zh": {
        # 侧边栏
        "lang_label": "Language",
        "lang_zh": "中文",
        "lang_en": "English",
        "system_status": "系统状态",
        "pro_version": "✅ 正式版",
        "trial_version": "⚠️ 试用版",
        "license_valid_until": "📅 有效期至 {exp}",
        "trial_used": "📊 已使用 {used} 次 / 共 {total} 次",
        "trial_exhausted": "❌ 试用次数已用完",
        "activate_title": "🔑 授权 / 续费",
        "activate_code_hint": "激活码",
        "activate_btn": "立即激活",
        "activate_success": "✅ 激活成功，有效期一年",
        "activate_fail": "❌ 激活码无效",
        "history_title": "📜 历史记录",
        
        # 主界面
        "main_title": "📊 8D报告智能生成助手",
        "input_header": "📝 输入基本信息",
        "product_name": "产品型号 / 名称",
        "customer": "客户名称",
        "problem_desc": "不良现象描述",
        "problem_placeholder": "请使用 5W2H 方法描述问题，例如：\n\n2024年3月15日，客户反馈批次号 240301 的电路板中，电阻 R12 存在虚焊现象，不良率约 0.5%（共2000件，发现10件）。",
        "occur_date": "发现日期",
        "defect_qty": "不良数量",
        "severity": "严重程度",
        "severity_low": "低",
        "severity_medium": "中",
        "severity_high": "高",
        "severity_critical": "危急",
        "industry_std": "适用标准",
        "team_members": "团队成员（可选）",
        "team_placeholder": "例：张明(组长), 李华(工程), 王芳(质量)",
        "generate_btn": "🚀 生成 8D 报告",
        "generating": "8D报告正在生成中...约60秒左右完成，请耐心等待。",
        
        # 预览区
        "preview_header": "📄 报告预览",
        "download_btn": "📥 导出 Word 报告",
        "export_disabled": "🔒 激活正式版后可导出 Word",
        
        # 提示信息
        "no_desc": "❌ 请输入不良现象描述",
        "trial_exhausted_error": "❌ 试用次数已用完，请购买激活码",
        "api_error": "❌ 服务异常，请稍后重试",
        "success": "✅ 报告生成完成！",
        
        # Word标题
        "word_title": "8D 问题纠正与预防措施报告"
    },
    "en": {
        # Sidebar
        "lang_label": "Language",
        "lang_zh": "中文",
        "lang_en": "English",
        "system_status": "System Status",
        "pro_version": "✅ Pro Version",
        "trial_version": "⚠️ Trial Version",
        "license_valid_until": "📅 Valid until {exp}",
        "trial_used": "📊 Used {used} / {total} times",
        "trial_exhausted": "❌ Trial exhausted",
        "activate_title": "🔑 License / Renew",
        "activate_code_hint": "Activation Code",
        "activate_btn": "Activate",
        "activate_success": "✅ Activated successfully, valid for one year",
        "activate_fail": "❌ Invalid activation code",
        "history_title": "📜 History",
        
        # Main Interface
        "main_title": "📊 8D Report Generator",
        "input_header": "📝 Input Information",
        "product_name": "Product Name / Model",
        "customer": "Customer Name",
        "problem_desc": "Problem Description",
        "problem_placeholder": "Please use 5W2H method, e.g.:\n\nOn March 15, 2024, customer reported soldering defect on resistor R12 in batch 240301, defect rate ~0.5% (10 out of 2000).",
        "occur_date": "Occurrence Date",
        "defect_qty": "Defect Quantity",
        "severity": "Severity",
        "severity_low": "Low",
        "severity_medium": "Medium",
        "severity_high": "High",
        "severity_critical": "Critical",
        "industry_std": "Standard",
        "team_members": "Team Members (Optional)",
        "team_placeholder": "e.g., Zhang(Leader), Li(Eng), Wang(Quality)",
        "generate_btn": "🚀 Generate 8D Report",
        "generating": "Generating 8D report... Please wait about 60 seconds.",
        
        # Preview
        "preview_header": "📄 Report Preview",
        "download_btn": "📥 Export Word Report",
        "export_disabled": "🔒 Activate to export Word",
        
        # Messages
        "no_desc": "❌ Please enter problem description",
        "trial_exhausted_error": "❌ Trial exhausted, please purchase license",
        "api_error": "❌ Service error, please try again later",
        "success": "✅ Report generated successfully!",
        
        # Word Title
        "word_title": "8D Corrective and Preventive Action Report"
    }
}

# =========================================================
# 2. 双语 System Prompt（根据语言切换）
# =========================================================
SYSTEM_PROMPT = {
    "zh": """
你是一位拥有20年经验的汽车电子行业高级质量工程师（SQE/CQE），精通 IATF 16949 标准和 8D 问题解决方法。
请根据用户输入的问题描述，撰写一份专业、逻辑严密的 8D 报告。

### 核心要求：
1. **语气**：专业、客观、工程化。拒绝空话，多用数据支撑。
2. **4M1E 分析 (D4)**：必须包含"人、机、料、法、环"五个维度。注意：**必须列出"排查正常"的项**，不要只写异常，要体现完整的排查过程。
3. **根本原因分析**：必须包含完整的[发生原因 5-Why]和[逃出原因 5-Why]分析。
4. **措施格式**：使用 `[责任人 | 完成时间 | 状态]` 格式。
5. **数据要求**：关键数据（如日期、数量、CPK值）要具体。

### 输出格式要求：
- D1到D8的标题必须写为："D1: 团队成立"（标题和内容在同一行，冒号后加一个空格）
- 每个D章节之间用空行分隔
- D1团队成立中，每个成员及其职责应该独立成行
- D2问题描述中，每个W（What、Where、When、Who、Why、How many、How）应该独立成行
- D4根本原因分析中，每个维度（人、机、料、法、环）应该独立成行，格式为：
  人：
  （内容）
  
  机：
  （内容）
- 不要使用任何Markdown标记（如**、#等）

### 输出结构（必须完整包含 D1-D8）：

D1: 团队成立
（团队列表和职责，每个成员一行）

D2: 问题描述
What: （什么缺陷）
Where: （在哪里发现）
When: （何时发现）
Who: （谁发现）
Why: （为何被发现）
How many: （数量多少）
How: （如何发现）

D3: 临时围堵措施
（包含具体数量和日期）

D4: 根本原因分析
人：
（人员方面分析）

机：
（设备方面分析）

料：
（物料方面分析）

法：
（方法方面分析）

环：
（环境方面分析）

（5-Why分析）

D5: 永久纠正措施
（针对根本原因的措施，使用[责任人 | 完成时间 | 状态]格式）

D6: 措施验证
（提供具体验证数据）

D7: 预防措施
（横向展开、文件更新）

D8: 结案与团队激励
（简短总结）

请直接输出8D报告正文，不要包含任何开场白或结束语。不要使用任何Markdown标记（如**、#等）。
""",
    "en": """
You are a Senior Quality Engineer (SQE/CQE) with 20 years of experience in the automotive electronics industry, proficient in IATF 16949 standards and 8D problem-solving methodology.
Based on the user's problem description, please write a professional and logically rigorous 8D report.

### Core Requirements:
1. **Tone**: Professional, objective, engineering-oriented. Avoid empty words, use data to support.
2. **4M1E Analysis (D4)**: Must include all five dimensions: Man, Machine, Material, Method, Environment. **Must list items that were checked and found normal**, not just abnormalities. Show the complete investigation process.
3. **Root Cause Analysis**: Must include complete [Occurrence Reason 5-Why] and [Escape Reason 5-Why] analysis.
4. **Action Format**: Use `[Responsible Person | Completion Date | Status]` format.
5. **Data Requirements**: Key data (such as dates, quantities, CPK values) should be specific.

### Output Format Requirements:
- D1 to D8 titles must be written as: "D1: Team Formation" (title and content on the same line, add a space after the colon)
- Separate each D section with a blank line
- In D1 Team Formation, each member and their responsibilities should be on its own line
- In D2 Problem Description, each W (What, Where, When, Who, Why, How many, How) should be on its own line
- In D4 Root Cause Analysis, each dimension (Man, Machine, Material, Method, Environment) should be on its own line, formatted as:
  Man:
  (content)
  
  Machine:
  (content)
- Do not use any Markdown markers (such as **, #, etc.)

### Output Structure (must include D1-D8 completely):

D1: Team Formation
(Team list and responsibilities, each member on its own line)

D2: Problem Description
What: (What defect)
Where: (Where found)
When: (When found)
Who: (Who found)
Why: (Why discovered)
How many: (Quantity)
How: (How discovered)

D3: Interim Containment Actions
(Include specific quantities and dates)

D4: Root Cause Analysis
Man:
(Personnel analysis)

Machine:
(Equipment analysis)

Material:
(Material analysis)

Method:
(Process method analysis)

Environment:
(Environmental analysis)

(5-Why analysis)

D5: Permanent Corrective Actions
(Actions targeting root causes, use [Responsible Person | Completion Date | Status] format)

D6: Action Verification
(Provide specific verification data)

D7: Preventive Actions
(Horizontal deployment, document updates)

D8: Closure and Team Recognition
(Brief summary)

Please output the 8D report content directly without any opening remarks or closing statements. Do not use any Markdown markers (such as **, #, etc.).
"""
}

# =========================================================
# 3. 内置 API 配置（用户不可见）
# =========================================================
try:
    API_KEY = st.secrets["DEEPSEEK_API_KEY"]
    BASE_URL = st.secrets["DEEPSEEK_BASE_URL"]
except:
    API_KEY = "your-api-key-here"
    BASE_URL = "https://api.deepseek.com"

MODEL = "deepseek-chat"

# =========================================================
# 4. License 本地持久化
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

def is_license_valid():
    exp = license_data.get("license_expire")
    if not exp:
        return False
    return datetime.today().date() <= datetime.strptime(exp, "%Y-%m-%d").date()

def activate_license():
    exp_date = datetime.today().date() + timedelta(days=LICENSE_DAYS)
    license_data["license_expire"] = exp_date.strftime("%Y-%m-%d")
    save_license(license_data)

def inc_trial_used():
    license_data["trial_used"] = license_data.get("trial_used", 0) + 1
    save_license(license_data)

# =========================================================
# 5. 输出后处理函数（增强 D1、D2、D4 换行处理）
# =========================================================
def clean_format(text):
    if not text:
        return ""
    
    # 1. 去除所有星号
    text = text.replace("**", "")
    # 2. 去除井号
    text = text.replace("#", "")
    
    # 3. 修复 Dx 标题换行问题：确保 "D1: 团队成立" 在同一行
    for i in range(1, 9):
        pattern = rf'(D{i}[:：])\s*\n+\s*'
        text = re.sub(pattern, rf'\1 ', text)
    
    # =========================================================
    # 4. D1 团队成员换行处理
    # =========================================================
    # 中文格式：张明(组长), 李华(工程), 王芳(质量)
    text = re.sub(r'([^，,]+[（(][^）)]+[）)])\s*[，,]\s*', r'\1\n', text)
    text = re.sub(r'([^，,]+[（(][^）)]+[）)])\s*(?=\n|$)', r'\1\n', text)
    # 英文格式：Zhang(Leader), Li(Eng), Wang(Quality)
    text = re.sub(r'([^,]+\([^)]+\))\s*,\s*', r'\1\n', text)
    text = re.sub(r'([^,]+\([^)]+\))\s*(?=\n|$)', r'\1\n', text)
    
    # =========================================================
    # 5. D2 5W2H 换行处理
    # =========================================================
    all_5w2h = ["What:", "Where:", "When:", "Who:", "Why:", "How many:", "How:"]
    for kw in all_5w2h:
        # 确保每个W前面有换行（如果不是行首）
        text = re.sub(rf'([^\n])({kw})', r'\1\n\2', text)
        # 确保W标题后换行
        text = re.sub(rf'({kw})([^\n])', r'\1\n\2', text)
    
    # =========================================================
    # 6. D4 4M1E 换行处理（关键修复）
    # =========================================================
    # 中文 4M1E 关键词
    chinese_4m1e = ["人：", "机：", "料：", "法：", "环："]
    for kw in chinese_4m1e:
        # 确保每个维度标题前面有换行（如果不是行首）
        text = re.sub(rf'([^\n])({kw})', r'\1\n\n\2', text)
        # 确保维度标题后换行（标题和内容分开）
        text = re.sub(rf'({kw})([^\n])', r'\1\n\2', text)
        # 处理内容中可能包含的"排查："等，确保格式整洁
        text = re.sub(rf'({kw}\n)([^人机料法环])', r'\1\2', text)
    
    # 英文 4M1E 关键词
    english_4m1e = ["Man:", "Machine:", "Material:", "Method:", "Environment:"]
    for kw in english_4m1e:
        text = re.sub(rf'([^\n])({kw})', r'\1\n\n\2', text)
        text = re.sub(rf'({kw})([^\n])', r'\1\n\2', text)
    
    # =========================================================
    # 7. 确保每个D章节之间有明确的分隔
    # =========================================================
    for i in range(1, 8):
        pattern = rf'(D{i}[:：][^\n]+\n)(?!\n)'
        text = re.sub(pattern, r'\1\n', text)
    
    # 8. 清理多余的空行（连续多个空行替换为两个空行）
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # 9. 处理 D4 中可能出现的连续换行问题
    # 确保每个维度之间有空行分隔
    for kw in chinese_4m1e:
        text = re.sub(rf'({kw}\n[^\n]+)\n+({kw})', r'\1\n\n\2', text)
    
    return text.strip()

# =========================================================
# 6. Word 导出功能（彻底去除星号，每个D之间用分割线分开）
# =========================================================
def export_to_word(content, product_name, lang):
    doc = Document()
    
    # 设置字体
    if lang == "zh":
        doc.styles['Normal'].font.name = '宋体'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    else:
        doc.styles['Normal'].font.name = 'Arial'
    
    doc.styles['Normal'].font.size = Pt(10.5)
    
    # 标题
    title = doc.add_heading(TEXT[lang]["word_title"], 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 产品信息
    info = doc.add_paragraph()
    info.add_run(f"Product: {product_name}").bold = True
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    # 彻底去除所有星号和井号
    clean_content = content.replace("**", "").replace("#", "")
    
    # 按 D1-D8 分割内容
    sections = []
    current_section = []
    lines = clean_content.split('\n')
    
    for line in lines:
        if re.match(r'^D[1-8][:：]', line.strip()):
            if current_section:
                sections.append('\n'.join(current_section))
                current_section = []
            current_section.append(line.strip())
        else:
            if current_section or line.strip():
                current_section.append(line)
    
    if current_section:
        sections.append('\n'.join(current_section))
    
    if len(sections) < 3:
        sections = re.split(r'\n(?=D[1-8][:：])', clean_content)
    
    for i, sec in enumerate(sections):
        if not sec.strip():
            continue
        
        lines = sec.strip().split('\n', 1)
        
        if not lines:
            continue
        
        title_text = lines[0].strip()
        if not re.match(r'^D[1-8][:：]', title_text):
            match = re.search(r'(D[1-8][:：][^\\n]*)', title_text)
            if match:
                title_text = match.group(1)
        
        p_title = doc.add_paragraph()
        runner = p_title.add_run(title_text)
        runner.bold = True
        runner.font.size = Pt(14)
        runner.font.color.rgb = RGBColor(30, 58, 138)
        
        if len(lines) > 1:
            content_text = lines[1].strip()
            if content_text:
                content_text = content_text.replace("**", "").replace("#", "")
                p_content = doc.add_paragraph(content_text)
                p_content.paragraph_format.line_spacing = 1.5
                p_content.paragraph_format.space_after = Pt(6)
        
        if i < len(sections) - 1:
            p_line = doc.add_paragraph()
            p_line.paragraph_format.space_before = Pt(12)
            p_line.paragraph_format.space_after = Pt(12)
            p = p_line._element
            pPr = p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '8')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), 'AAAAAA')
            pBdr.append(bottom)
            pPr.append(pBdr)
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =========================================================
# 7. Session State 初始化
# =========================================================
if "lang" not in st.session_state:
    st.session_state.lang = "zh"
if "current_result" not in st.session_state:
    st.session_state.current_result = ""
if "history" not in st.session_state:
    st.session_state.history = []

T = TEXT[st.session_state.lang]

# =========================================================
# 8. 侧边栏（语言 + 授权 + 历史）
# =========================================================
with st.sidebar:
    # 语言切换
    st.radio(
        T["lang_label"],
        [T["lang_zh"], T["lang_en"]],
        index=0 if st.session_state.lang == "zh" else 1,
        horizontal=True,
        key="lang_radio"
    )
    if st.session_state.lang_radio != T["lang_zh"] and st.session_state.lang != "en":
        st.session_state.lang = "en"
        T = TEXT[st.session_state.lang]
        st.rerun()
    elif st.session_state.lang_radio != T["lang_en"] and st.session_state.lang != "zh":
        st.session_state.lang = "zh"
        T = TEXT[st.session_state.lang]
        st.rerun()
    
    st.markdown("---")
    
    # 系统状态
    st.header(T["system_status"])
    
    is_valid = is_license_valid()
    exp = license_data.get("license_expire")
    
    if is_valid:
        st.success(T["pro_version"])
        if exp:
            st.caption(T["license_valid_until"].format(exp=exp))
    else:
        st.warning(T["trial_version"])
        used = license_data.get("trial_used", 0)
        rem = trial_remaining()
        st.caption(T["trial_used"].format(used=used, total=MAX_TRIAL_GENERATIONS))
        if rem <= 0:
            st.error(T["trial_exhausted"])
        if exp:
            st.caption(T["license_valid_until"].format(exp=exp))
    
    st.markdown("---")
    
    # 授权激活
    st.subheader(T["activate_title"])
    code = st.text_input(T["activate_code_hint"], type="password")
    
    if st.button(T["activate_btn"], use_container_width=True):
        if len(code) >= 8:
            activate_license()
            st.success(T["activate_success"])
            st.rerun()
        else:
            st.error(T["activate_fail"])
    
    st.markdown("---")
    
    # 历史记录
    st.header(T["history_title"])
    
    if st.session_state.history:
        for i, item in enumerate(st.session_state.history[-5:]):
            if st.button(f"📄 {item['title']}", key=f"h_{i}", use_container_width=True):
                st.session_state.current_result = item["content"]
                st.rerun()
    else:
        st.caption("暂无历史记录" if st.session_state.lang == "zh" else "No history")

# =========================================================
# 9. 主界面
# =========================================================
st.title(T["main_title"])
st.markdown("---")

col_input, col_preview = st.columns([1, 1.2])

with col_input:
    st.header(T["input_header"])
    
    # 基本信息
    product_name = st.text_input(T["product_name"], placeholder="例：PCB-A123")
    customer = st.text_input(T["customer"], placeholder="例：比亚迪汽车")
    
    # 问题描述
    problem_desc = st.text_area(
        T["problem_desc"],
        height=150,
        placeholder=T["problem_placeholder"]
    )
    
    # 第二行信息
    col1, col2, col3 = st.columns(3)
    with col1:
        occur_date = st.date_input(T["occur_date"], datetime.now())
    with col2:
        defect_qty = st.number_input(T["defect_qty"], min_value=1, value=1, step=1)
    with col3:
        severity = st.selectbox(
            T["severity"],
            [T["severity_low"], T["severity_medium"], T["severity_high"], T["severity_critical"]]
        )
    
    # 第三行信息
    col4, col5 = st.columns(2)
    with col4:
        industry_std = st.selectbox(
            T["industry_std"],
            ["ISO 9001 (General)", "IATF 16949 (Automotive)", "ISO 13485 (Medical)", "AS9100 (Aerospace)"],
            index=0
        )
    with col5:
        team_members = st.text_input(
            T["team_members"],
            placeholder=T["team_placeholder"]
        )
    
    st.markdown("---")
    
    # 生成按钮
    if st.button(T["generate_btn"], type="primary", use_container_width=True):
        is_valid = is_license_valid()
        
        if not is_valid and trial_remaining() <= 0:
            st.error(T["trial_exhausted_error"])
            st.stop()
        
        if not problem_desc:
            st.error(T["no_desc"])
        else:
            with st.spinner(T["generating"]):
                try:
                    client = openai.OpenAI(
                        api_key=API_KEY,
                        base_url=BASE_URL
                    )
                    
                    # 根据当前语言选择对应的 System Prompt
                    current_lang = st.session_state.lang
                    system_prompt = SYSTEM_PROMPT[current_lang]
                    
                    severity_map = {
                        T["severity_low"]: "Low",
                        T["severity_medium"]: "Medium",
                        T["severity_high"]: "High",
                        T["severity_critical"]: "Critical"
                    }
                    
                    # 根据语言组装用户提示词
                    if current_lang == "zh":
                        user_prompt = f"""
请根据以下信息生成完整的 8D 报告：

产品信息：
- 产品名称/型号：{product_name or '未提供'}
- 客户名称：{customer or '未提供'}
- 发现日期：{occur_date}
- 不良数量：{defect_qty}
- 严重程度：{severity_map[severity]}
- 适用标准：{industry_std}
- 团队成员：{team_members or '未提供'}

问题描述：
{problem_desc}

请严格按照系统提示词的要求，输出 D1 到 D8 的完整报告。
"""
                    else:
                        user_prompt = f"""
Please generate a complete 8D report based on the following information:

Product Information:
- Product Name/Model: {product_name or 'Not provided'}
- Customer Name: {customer or 'Not provided'}
- Occurrence Date: {occur_date}
- Defect Quantity: {defect_qty}
- Severity: {severity_map[severity]}
- Applicable Standard: {industry_std}
- Team Members: {team_members or 'Not provided'}

Problem Description:
{problem_desc}

Please output the complete D1 to D8 report according to the system prompt requirements.
"""
                    
                    response = client.chat.completions.create(
                        model=MODEL,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_prompt}
                        ],
                        temperature=0.2,
                        timeout=90
                    )
                    
                    raw_result = response.choices[0].message.content
                    final_result = clean_format(raw_result)
                    
                    st.session_state.current_result = final_result
                    
                    st.session_state.history.append({
                        "title": f"{product_name or '8D'}_{occur_date.strftime('%Y%m%d')}",
                        "content": final_result
                    })
                    
                    if not is_valid:
                        inc_trial_used()
                    
                    st.success(T["success"])
                    st.rerun()
                    
                except Exception as e:
                    st.error(T["api_error"])

with col_preview:
    st.header(T["preview_header"])
    
    if st.session_state.current_result:
        # 预览时也去除星号和井号
        preview_content = st.session_state.current_result.replace("**", "").replace("#", "")
        st.markdown(preview_content)
        
        st.markdown("---")
        
        if is_license_valid():
            word_data = export_to_word(
                st.session_state.current_result,
                product_name or "8D_Report",
                st.session_state.lang
            )
            st.download_button(
                label=T["download_btn"],
                data=word_data,
                file_name=f"8D_Report_{product_name or 'Report'}_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.info(T["export_disabled"])
    else:
        st.info("👈 输入问题描述后点击生成按钮" if st.session_state.lang == "zh" else "👈 Enter problem description and click generate")
