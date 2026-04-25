#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
8D 报告智能生成助手 - 客户端
最终优化版本
"""

import streamlit as st
from io import BytesIO
from datetime import datetime, timedelta
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import openai
from supabase import create_client
import logging
import time

# ==================== 缓存配置 ====================
@st.cache_data(ttl=60)
def get_cached_license(user_id):
    """缓存用户许可证信息，60 秒 TTL"""
    if not supabase:
        return None
    try:
        r = supabase.table("licenses").select("*").eq("user_id", user_id).execute()
        if r.data:
            return r.data[0]
        return create_free_license(user_id)
    except Exception:
        return None

def clear_license_cache(user_id):
    """清除特定用户的缓存"""
    get_cached_license.clear()

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="8D 报告 - 智能生成助手", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 隐藏 Streamlit 默认 UI 元素 ====================
# ==================== 隐藏 Streamlit 默认 UI 元素 ====================
hide_streamlit_style = """
<style>
    /* 只隐藏右上角菜单 */
    #MainMenu {visibility: hidden !important; display: none !important;}
    
    /* 隐藏 footer 水印 */
    footer {visibility: hidden !important; display: none !important;}
    
    /* 隐藏 Pages 导航菜单列表 */
    [data-testid="stSidebarNav"] > ul {display: none !important;}
    
    /* 调整主内容区域 */
    .main .block-container {
        padding-top: 0.5rem !important;
    }
    
    /* ========== 缩小侧边栏间距 ========== */
    [data-testid="stSidebar"] .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
    [data-testid="stSidebar"] h3 {
        margin-top: 0.5rem !important;
        margin-bottom: 0.3rem !important;
    }
    
    [data-testid="stSidebar"] p {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    [data-testid="stSidebar"] .stButton {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    [data-testid="stSidebar"] .stTextInput {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    [data-testid="stSidebar"] .streamlit-expanderHeader {
        padding-top: 0.3rem !important;
        padding-bottom: 0.3rem !important;
    }
    
    [data-testid="stSidebar"] .stCaption {
        margin-top: 0.1rem !important;
        margin-bottom: 0.1rem !important;
    }
    
    [data-testid="stSidebar"] .stMarkdown {
        margin-bottom: 0.3rem !important;
    }
    
    [data-testid="stSidebar"] div[data-testid="stVerticalBlock"] > div {
        gap: 0.15rem !important;
    }
    
    [data-testid="stSidebar"] hr {
        display: none !important;
    }
    
    /* ========== 手机端字体缩小 ========== */
    @media screen and (max-width: 768px) {
        h1 { font-size: 1.3rem !important; }
        h2 { font-size: 1.1rem !important; }
        h3 { font-size: 1rem !important; }
        h4 { font-size: 0.95rem !important; }
        body { font-size: 0.85rem !important; }
        input, textarea { font-size: 0.85rem !important; }
        button { font-size: 0.9rem !important; }
        label, .stMarkdown p, .stMarkdown span { font-size: 0.85rem !important; }
                /* 手机端缩小文本输入框高度 */
        textarea {
            height: 80px !important;
            min-height: 80px !important;
    }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==================== JavaScript 隐藏右上角按钮 ====================
hide_buttons_script = """
<script>
// Hide top-right toolbar buttons, keep sidebar toggle
function hideTopRightButtons() {
    // Hide links containing fork, github, star
    document.querySelectorAll('a').forEach(el => {
        const href = (el.href || '').toLowerCase();
        const text = (el.textContent || '').toLowerCase();
        if (href.includes('fork') || href.includes('github') || 
            text.includes('fork') || text.includes('star')) {
            el.style.display = 'none';
        }
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(hideTopRightButtons, 500));
} else {
    setTimeout(hideTopRightButtons, 500);
}
</script>
"""
st.markdown(hide_buttons_script, unsafe_allow_html=True)

# ==================== 多语言文本 ====================
TEXT = {
    "zh": {
        "lang_label": "语言", "lang_zh": "中文", "lang_en": "English",
        "system_status": "系统状态", "pro_version": "✅ 正式版", "trial_version": "⚠️ 试用版",
        "license_valid_until": "📅 有效期至 {exp}", "trial_used": "📊 已使用 {used} 次 / 共 {total} 次",
        "trial_exhausted": "❌ 试用次数已用完", "activate_title": "🔑 授权 / 续费",
        "activate_code_hint": "激活码", "activate_btn": "立即激活",
        "activate_success": "✅ 激活成功，有效期一年", "activate_fail": "❌ 激活码无效",
        "invalid_activate_code": "请输入有效的激活码",
        "license_expired": "❌ 授权已过期",
        "login_required": "🔒 请先登录", "logout": "退出登录",
        "login_header": "👤 用户登录",
        "username_label": "邮箱或手机号",
        "username_placeholder": "例：zhangsan@163.com 或 13812345678",
        "login_register_btn": "🔓 登录 / 注册",
        "enter_username_error": "请输入邮箱或手机号",
        "invalid_email": "❌ 邮箱格式不正确，示例：zhangsan@163.com",
        "invalid_phone": "❌ 手机号格式不正确，请输入11位大陆手机号（1开头，第二位3-9）",
        "invalid_contact": "❌ 请输入有效的邮箱或11位大陆手机号",
        "expander_activate_code": "🔑 输入激活码",
        "enter_activate_code_placeholder": "输入激活码",
        "trial_remaining": "📊 **试用版** | 剩余 {n} 次",
        "valid_until": "⏰ 有效期至: {date}",
        "valid_until_date": "📅 有效期至: {date}",
        "permanent_valid": "♾️ 永久有效",
        "account_manager": "🔐 账户管理 / Account",
        "contact_service": "📱 联系客服 / Contact",
        "new_user_hint": "👋 新用户登录赠送3次免费试用！",
        "main_title": "📊 8D 报告智能生成助手",
        "progress_phases": [
            {"icon": "📝", "text": "正在整理您的输入信息...", "sub": "产品：{product}"},
            {"icon": "🤔", "text": "正在理解问题背景...", "sub": "运用 5W2H 方法分析"},
            {"icon": "📋", "text": "正在生成 D1 团队组建...", "sub": "确定责任人及时间节点"},
            {"icon": "📋", "text": "正在生成 D2 问题描述...", "sub": "详细记录不良现象"},
            {"icon": "🛡️", "text": "正在生成 D3 临时措施...", "sub": "遏制问题扩散"},
            {"icon": "🔬", "text": "正在分析根本原因 (4M1E)...", "sub": "人、机、料、法、环逐一排查"},
            {"icon": "🔍", "text": "正在进行 5-Why 追问...", "sub": "追溯至根本原因"},
            {"icon": "💡", "text": "正在制定 D5 永久措施...", "sub": "根本性解决方案"},
            {"icon": "✅", "text": "正在生成 D6 实施计划...", "sub": "验证措施有效性"},
            {"icon": "📊", "text": "正在生成 D7 预防措施...", "sub": "防止问题复发"},
            {"icon": "🏆", "text": "正在生成 D8 总结表彰...", "sub": "固化经验，分享成果"},
            {"icon": "✨", "text": "正在优化报告格式...", "sub": "确保专业美观"},
        ], "input_header": "📝 输入基本信息",
        "product_name": "产品型号 / 名称", "customer": "客户名称",
        "problem_desc": "不良现象描述",
        "problem_placeholder": "请使用 5W2H 方法描述问题",
        "occur_date": "发现日期", "defect_qty": "不良数量", "severity": "严重程度",
        "severity_low": "低", "severity_medium": "中", "severity_high": "高", "severity_critical": "危急",
        "industry_std": "适用标准", "team_members": "团队成员（可选）",
        "team_placeholder": "例：张明 (组长), 李华 (工程)",
        "generate_btn": "🚀 生成 8D 报告", 
        "generating": "8D 报告智能生成中，请稍候...",
        "preview_header": "📄 报告预览", "download_btn": "📥 导出 Word 报告",
        "export_disabled": "🔒 激活正式版后可导出 Word",
        "no_desc": "❌ 请输入不良现象描述",
        "trial_exhausted_error": "❌ 试用次数已用完", "api_error": "❌ 服务异常",
        "success": "✅ 报告生成完成！", "report_complete": "报告生成完成！",
        "beautifying": "正在美化格式...", "word_title": "8D 问题纠正与预防措施报告",
        "system_error": "❌ 系统错误，请稍后重试",
    },
    "en": {
        "lang_label": "Language", "lang_zh": "中文", "lang_en": "English",
        "system_status": "System Status", "pro_version": "✅ Pro Version", "trial_version": "⚠️ Trial Version",
        "license_valid_until": "📅 Valid until {exp}", "trial_used": "📊 Used {used} / {total}",
        "trial_exhausted": "❌ Trial exhausted", "activate_title": "🔑 License / Renew",
        "activate_code_hint": "Activation Code", "activate_btn": "Activate",
        "activate_success": "✅ Activated successfully", "activate_fail": "❌ Invalid code",
        "invalid_activate_code": "Please enter a valid activation code",
        "license_expired": "❌ License expired",
        "login_required": "🔒 Please login", "logout": "Logout",
        "login_header": "👤 User Login",
        "username_label": "Email or Phone",
        "username_placeholder": "e.g., name@example.com or 13812345678",
        "login_register_btn": "🔓 Login / Register",
        "enter_username_error": "Please enter email or phone number",
        "invalid_email": "❌ Invalid email format, e.g., name@example.com",
        "invalid_phone": "❌ Invalid phone number, enter 11-digit mainland China number",
        "invalid_contact": "❌ Please enter a valid email or 11-digit phone number",
        "expander_activate_code": "🔑 Enter Activation Code",
        "enter_activate_code_placeholder": "Enter activation code",
        "trial_remaining": "📊 **Trial** | {n} remaining",
        "valid_until": "⏰ Valid until: {date}",
        "valid_until_date": "📅 Valid until: {date}",
        "permanent_valid": "♾️ Permanent",
        "account_manager": "🔐 Account Manager",
        "contact_service": "📱 Contact Service",
        "new_user_hint": "👋 New user? Enter email or phone to register, get 3 free trials!",
        "main_title": "📊 8D Report Generator",
        "progress_phases": [
            {"icon": "📝", "text": "Organizing your input...", "sub": "Product: {product}"},
            {"icon": "🤔", "text": "Analyzing context...", "sub": "Using 5W2H method"},
            {"icon": "📋", "text": "Generating D1 Team...", "sub": "Defining responsibilities"},
            {"icon": "📋", "text": "Generating D2 Description...", "sub": "Recording defect details"},
            {"icon": "🛡️", "text": "Generating D3 Containment...", "sub": "Preventing spread"},
            {"icon": "🔬", "text": "Analyzing root cause (4M1E)...", "sub": "Checking all factors"},
            {"icon": "🔍", "text": "Performing 5-Why analysis...", "sub": "Finding root cause"},
            {"icon": "💡", "text": "Developing D5 Actions...", "sub": "Long-term solutions"},
            {"icon": "✅", "text": "Generating D6 Implementation...", "sub": "Verifying effectiveness"},
            {"icon": "📊", "text": "Generating D7 Prevention...", "sub": "Preventing recurrence"},
            {"icon": "🏆", "text": "Generating D8 Closure...", "sub": "Documenting lessons"},
            {"icon": "✨", "text": "Formatting report...", "sub": "Professional output"},
        ], "input_header": "📝 Input Information",
        "product_name": "Product Name / Model", "customer": "Customer Name",
        "problem_desc": "Problem Description",
        "problem_placeholder": "Please use 5W2H method",
        "occur_date": "Occurrence Date", "defect_qty": "Defect Quantity", "severity": "Severity",
        "severity_low": "Low", "severity_medium": "Medium", "severity_high": "High", "severity_critical": "Critical",
        "industry_std": "Standard", "team_members": "Team Members (Optional)",
        "team_placeholder": "e.g., Zhang(Leader), Li(Eng)",
        "generate_btn": "🚀 Generate 8D Report",
        "generating": "Generating report, please wait...",
        "preview_header": "📄 Report Preview", "download_btn": "📥 Export Word",
        "export_disabled": "🔒 Activate to export",
        "no_desc": "❌ Please enter description",
        "trial_exhausted_error": "❌ Trial exhausted", "api_error": "❌ Service error",
        "success": "✅ Report generated!", "report_complete": "Report generated!",
        "beautifying": "Formatting...", "word_title": "8D Corrective Action Report",
        "system_error": "❌ System error, please try again later",
    }
}

# ==================== 系统提示词 ====================
SYSTEM_PROMPT = {
    "zh": (
        "你是一位拥有 20 年经验的汽车电子行业高级质量工程师，精通 IATF 16949 标准和 8D 问题解决方法。"
        "请根据用户输入撰写专业、逻辑严密的 8D 报告。\n\n"
        "【D4 根本原因分析要求】\n"
        "4M1E 分析（人、机、料、法、环）要逐项确认，使用确定句而非疑问句：\n"
        "✅ 正常项：明确说明\"经检查，XX 符合标准，排除为根本原因\"\n"
        "❌ 异常项：明确说明\"经检查，XX 存在问题：[具体问题]\"\n"
        "❌ 不要使用\"是否\"、\"有没有\"等疑问句\n\n"
        "使用 5-Why 分析法：\n"
        "从异常项开始，连续追问\"为什么\"\n"
        "至少追问 3-5 层，直到找到根本原因\n"
        "每层回答要具体，不能笼统\n\n"
        "输出格式示例（注意换行）：\n"
        "【4M1E 分析】\n\n"
        "人：经检查，操作员持证上岗 → 排除\n\n"
        "机：经检查，设备参数偏移 0.05mm → 异常项 ⚠️\n\n"
        "料：经检查，原材料合格 → 排除\n\n"
        "法：经检查，作业指导书过期 → 异常项 ⚠️\n\n"
        "环：经检查，环境符合要求 → 排除\n\n"
        "【5-Why 分析】\n\n"
        "Why1：为什么设备参数偏移？→ 传感器校准过期\n\n"
        "Why2：为什么校准过期？→ 年度校准计划未执行\n\n"
        "Why3：为什么计划未执行？→ 维护人员不足\n\n"
        "Why4：为什么人员不足？→ 未配置备用人员\n\n"
        "Why5：为什么未配置备用人员？→ 人员编制申请未获批 ← 根本原因\n\n"
        "【其他要求】\n"
        "语气专业客观\n"
        "措施使用 [责任人 | 时间 | 状态] 格式\n"
        "不使用 Markdown 标记\n"
        "直接输出 D1-D8 报告正文\n"
        "4M1E 分析必须使用确定句，明确指出哪些正常、哪些异常"
    ),
    
    "en": (
        "You are a Senior Quality Engineer with 20 years experience in automotive electronics, "
        "proficient in IATF 16949 and 8D methodology. Please write a professional 8D report based on user input.\n\n"
        "【D4 Root Cause Analysis Requirements】\n"
        "4M1E analysis (Man, Machine, Material, Method, Environment) must use declarative sentences:\n"
        "✅ Normal items: Clearly state \"Verified, XX meets standard, excluded as root cause\"\n"
        "❌ Abnormal items: Clearly state \"Verified, XX has issue: [specific problem]\"\n"
        "❌ Do NOT use questions like \"whether\", \"is there\"\n\n"
        "Use 5-Why analysis:\n"
        "Start from abnormal items, continuously ask \"why\"\n"
        "At least 3-5 levels until finding root cause\n"
        "Each answer must be specific, not vague\n\n"
        "【Other Requirements】\n"
        "Professional tone\n"
        "Use [Owner|Date|Status] format for actions\n"
        "No Markdown\n"
        "Output D1-D8 directly"
    )
}
# ==================== 初始化配置 ====================
try:
    API_KEY = st.secrets["DEEPSEEK_API_KEY"]
    BASE_URL = st.secrets["DEEPSEEK_BASE_URL"]
except Exception:
    API_KEY = ""
    BASE_URL = "https://api.deepseek.com"

try:
    supabase = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
except Exception:
    supabase = None

# ==================== 核心功能函数 ====================
def get_user_license(user_id):
    return get_cached_license(user_id)

def create_free_license(user_id):
    if not supabase:
        return None
    try:
        r = supabase.table("licenses").insert({
            "user_id": user_id,
            "plan_type": "free",
            "trial_used": 0,
            "trial_limit": 3
        }).execute()
        return r.data[0] if r.data else None
    except Exception:
        return None

def can_generate_report(user_id):
    lic = get_user_license(user_id)
    if not lic:
        return False
    if lic['plan_type'] in ['free', 'trial']:
        return lic['trial_used'] < lic['trial_limit']
    if lic['plan_type'] in ['pro', 'enterprise']:
        if lic.get('license_expire'):
            try:
                return datetime.now() < datetime.fromisoformat(lic['license_expire'])
            except Exception:
                return True
        return True
    return False

def inc_trial_used(user_id):
    if not supabase:
        return
    try:
        lic = get_cached_license(user_id)
        if lic:
            new_count = lic.get('trial_used', 0) + 1
            supabase.table("licenses").update({"trial_used": new_count}).eq("user_id", user_id).execute()
            supabase.table("usage_logs").insert({
                "user_id": user_id,
                "action": "generate_report",
                "created_at": datetime.now().isoformat()
            }).execute()
            clear_license_cache(user_id)
    except Exception as e:
        logging.error(f"更新试用次数失败：{e}")

def activate_license_code(user_id, code):
    if not supabase:
        return False, "系统错误"
    try:
        r = supabase.table("activation_codes").select("*").eq("code", code.strip().upper()).execute()
        if not r.data:
            return False, "无效的激活码"
        ac = r.data[0]
        if ac.get('is_used'):
            return False, "激活码已被使用"
        if ac.get('expire_date'):
            if datetime.now().date() > datetime.fromisoformat(ac['expire_date']).date():
                return False, "激活码已过期"
        duration = ac.get('duration_days', 365)
        exp_date = (datetime.now() + timedelta(days=duration)).isoformat()
        supabase.table("licenses").upsert({
            "user_id": user_id,
            "plan_type": ac.get('plan_type', 'pro'),
            "license_expire": exp_date
        }, on_conflict="user_id").execute()
        supabase.table("activation_codes").update({
            "is_used": True,
            "used_by": user_id,
            "used_at": datetime.now().isoformat()
        }).eq("code", code.strip().upper()).execute()
        clear_license_cache(user_id)
        formatted_date = exp_date[:10] if len(exp_date) >= 10 else exp_date
        return True, f"激活成功！有效期至 {formatted_date}"
    except Exception as e:
        logging.error(f"激活失败：{e}")
        return False, f"激活失败：{str(e)}"

def clean_format(text):
    if not text:
        return ""
    text = text.replace("**", "").replace("#", "")
    for i in range(1, 9):
        text = re.sub(rf'(D{i}[:：])\s*\n+\s*', rf'\1 ', text)
    text = re.sub(r'([人机料法环]：)', r'\n\1', text)
    text = re.sub(r'(→ 排除|→ 异常项[^，]*？)', r'\1\n', text)
    text = re.sub(r'(Why\d+：)', r'\n\1', text)
    text = re.sub(r'(→ [^\n]+)(?=Why\d+：|$)', r'\1\n', text)
    text = re.sub(r'(为什么\d+：)', r'\n\1', text)
    text = re.sub(r'(→ [^\n]+)(?=为什么\d+：|$)', r'\1\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def export_to_word(content, product_name, lang):
    doc = Document()
    if lang == "zh":
        doc.styles['Normal'].font.name = '宋体'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    else:
        doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10.5)
    title = doc.add_heading(TEXT[lang]["word_title"], 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = doc.add_paragraph()
    info.add_run(f"Product: {product_name}").bold = True
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    sections = re.split(r'\n(?=D[1-8][:：])', content.replace("**", "").replace("#", ""))
    for i, sec in enumerate(sections):
        if not sec.strip():
            continue
        lines = sec.strip().split('\n', 1)
        p_title = doc.add_paragraph()
        runner = p_title.add_run(lines[0].strip())
        runner.bold = True
        runner.font.size = Pt(14)
        runner.font.color.rgb = RGBColor(30, 58, 138)
        if len(lines) > 1 and lines[1].strip():
            doc.add_paragraph(lines[1].strip())
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
            pBdr.append(bottom)
            pPr.append(pBdr)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==================== 会话状态初始化 ====================
if "lang" not in st.session_state:
    st.session_state.lang = "zh"
if "current_result" not in st.session_state:
    st.session_state.current_result = ""
if "user_id" not in st.session_state:
    st.session_state.user_id = None

T = TEXT[st.session_state.lang]

# ==================== 侧边栏（登录和用户管理） ====================
def render_sidebar():
    """渲染侧边栏 - 登录、用户信息、语言切换、激活码等"""
    T = TEXT[st.session_state.lang]
    
    with st.sidebar:
        # ==================== 语言切换 ====================
        st.markdown("### 🌐 语言 / Language")
        lang_option = st.selectbox(
            "选择语言 / Select Language",
            ["中文", "English"],
            index=0 if st.session_state.lang == "zh" else 1,
            key="sidebar_lang_select",
            label_visibility="collapsed"
        )
        new_lang = "zh" if lang_option == "中文" else "en"
        if new_lang != st.session_state.lang:
            st.session_state.lang = new_lang
            st.rerun()
        
        st.markdown(f"### {T['account_manager']}")
                
        # ==================== 登录 / 注册区域 ====================
        user_id = st.session_state.get("user_id")
        
        if not user_id:
            st.info(T["new_user_hint"])
            
            user_input = st.text_input(
                T["username_label"],
                key="sidebar_user_input",
                placeholder=T["username_placeholder"]
            )
            
            # ========== 格式校验函数 ==========
            def validate_contact(contact):
                """校验邮箱或手机号格式"""
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                phone_pattern = r'^1[3-9]\d{9}$'
                
                if re.match(email_pattern, contact):
                    return True, "email"
                elif re.match(phone_pattern, contact):
                    return True, "phone"
                else:
                    return False, None
            
            if st.button(T["login_register_btn"], use_container_width=True, key="sidebar_login_btn"):
                if not user_input:
                    st.error(T["enter_username_error"])
                    st.stop()
                
                # 检查是否为已注册的老用户
                existing_user = False
                if supabase:
                    try:
                        existing = supabase.table("licenses").select("user_id").eq("user_id", user_input).execute()
                        existing_user = existing.data is not None and len(existing.data) > 0
                    except Exception:
                        st.error(T["system_error"])
                        st.stop()
                
                # 校验格式
                is_valid, contact_type = validate_contact(user_input)
                
                # 老用户不受格式限制，直接放行
                if not existing_user and not is_valid:
                    if "@" in user_input:
                        st.error(T["invalid_email"])
                    elif user_input.startswith("1") and len(user_input) == 11:
                        st.error(T["invalid_phone"])
                    else:
                        st.error(T["invalid_contact"])
                    st.stop()
                
                # 新用户注册（格式验证通过 且 无历史记录）
                if not existing_user and supabase:
                    try:
                        supabase.table("licenses").insert({
                            "user_id": user_input,
                            "plan_type": "free",
                            "trial_used": 0,
                            "trial_limit": 3
                        }).execute()
                    except Exception:
                        pass
                
                st.session_state.user_id = user_input
                st.rerun()
        
        # ==================== 已登录用户区域 ====================
        else:
            lic = get_user_license(user_id)
            
            st.markdown(f"**👤 {user_id[:30]}**")
            
            if lic:
                if lic.get('plan_type') == 'free':
                    remaining = lic['trial_limit'] - lic['trial_used']
                    st.info(T["trial_remaining"].format(n=remaining))
                    if lic.get('license_expire'):
                        try:
                            exp_date = datetime.fromisoformat(lic['license_expire']).strftime('%Y-%m-%d')
                            st.caption(T["valid_until"].format(date=exp_date))
                        except:
                            pass
                else:
                    st.success(T["pro_version"])
                    if lic.get('license_expire'):
                        try:
                            exp_date = datetime.fromisoformat(lic['license_expire']).strftime('%Y-%m-%d')
                            st.caption(T["valid_until_date"].format(date=exp_date))
                        except:
                            pass
                    else:
                        st.caption(T["permanent_valid"])
            
                        
            with st.expander(T["expander_activate_code"], expanded=False):
                activate_code = st.text_input(
                    T["activate_code_hint"],
                    type="password",
                    key="sidebar_act_code",
                    placeholder=T["enter_activate_code_placeholder"]
                )
                if st.button(T["activate_btn"], key="sidebar_act_btn", use_container_width=True):
                    if activate_code and len(activate_code) >= 6:
                        success, msg = activate_license_code(user_id, activate_code)
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.error(T["invalid_activate_code"])
            
            if st.button(T["logout"], key="sidebar_logout_btn", use_container_width=True):
                st.session_state.user_id = None
                st.session_state.current_result = ""
                get_cached_license.clear()
                st.rerun()
        
        # ==================== 底部信息 ====================
        
        st.markdown(f"**{T['contact_service']}**")
        try:
            st.image("wechat_qrcode.jpg", width=180)
        except:
            st.info("微信二维码：907749064")
        st.caption("淘宝店铺: 效率工坊铺")
        st.caption("微信号Wechat: 907749064")
        st.caption("Email: 907749064@qq.com")
        st.markdown("---")
 
# ==================== 主页面 ====================
render_sidebar()

st.title(T["main_title"])

col_input, col_preview = st.columns([1, 1.2])

with col_input:
    st.header(T["input_header"])
    product_name = st.text_input(T["product_name"], placeholder="e.g., PCB-A123" if st.session_state.lang == "en" else "例：PCB-A123")
    customer = st.text_input(T["customer"], placeholder="e.g., BYD" if st.session_state.lang == "en" else "例：比亚迪汽车")
    problem_desc = st.text_area(T["problem_desc"], height=150, placeholder=T["problem_placeholder"])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        occur_date = st.date_input(T["occur_date"], datetime.now())
    with col2:
        defect_qty = st.number_input(T["defect_qty"], min_value=1, value=1)
    with col3:
        severity = st.selectbox(
            T["severity"], 
            [T["severity_low"], T["severity_medium"], T["severity_high"], T["severity_critical"]]
        )
    
    col4, col5 = st.columns(2)
    with col4:
        industry_std = st.selectbox(
            T["industry_std"], 
            ["ISO 9001", "IATF 16949", "ISO 13485", "AS9100"], 
            index=1
        )
    with col5:
        team_members = st.text_input(T["team_members"], placeholder=T["team_placeholder"])
    
        
    if st.button(T["generate_btn"], type="primary", use_container_width=True):
        if not st.session_state.get("user_id"):
            st.error(T["login_required"])
            st.stop()
        
        user_id = st.session_state.user_id
        if not can_generate_report(user_id):
            lic = get_user_license(user_id)
            if lic and lic['plan_type'] == 'free':
                st.error(T["trial_exhausted_error"])
            else:
                st.error(T["license_expired"])
            st.stop()
        
        if not problem_desc:
            st.error(T["no_desc"])
        else:
            # 使用多语言进度条
            progress_phases = T["progress_phases"]
            # 替换产品名称占位符
            for phase in progress_phases:
                phase["sub"] = phase["sub"].replace("{product}", product_name or "未提供" if st.session_state.lang == "zh" else "Not provided")
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            sub_text = st.empty()
            
            with st.spinner(T["generating"]):
                try:
                    phase = progress_phases[0]
                    status_text.markdown(f"### {phase['icon']} {phase['text']}")
                    sub_text.caption(phase['sub'])
                    progress_bar.progress(1 / len(progress_phases))
                    
                    for i in range(1, len(progress_phases) - 2):
                        phase = progress_phases[i]
                        status_text.markdown(f"### {phase['icon']} {phase['text']}")
                        sub_text.caption(phase['sub'])
                        progress_bar.progress((i + 1) / len(progress_phases))
                        time.sleep(5)  # 每阶段停留约5秒

                    phase = progress_phases[-2]
                    status_text.markdown(f"### {phase['icon']} {phase['text']}")
                    sub_text.caption(phase['sub'])
                    progress_bar.progress(0.9)
                    
                    client = openai.OpenAI(api_key=API_KEY, base_url=BASE_URL)
                    user_prompt = (
                        f"请根据以下信息生成 8D 报告："
                        f"产品：{product_name or '未提供'}, "
                        f"客户：{customer or '未提供'}, "
                        f"日期：{occur_date}, "
                        f"数量：{defect_qty}, "
                        f"严重程度：{severity}, "
                        f"标准：{industry_std}, "
                        f"团队：{team_members or '未提供'}\n\n"
                        f"问题描述：{problem_desc}"
                    )
                    
                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": SYSTEM_PROMPT[st.session_state.lang]},
                            {"role": "user", "content": user_prompt}
                        ],
                        temperature=0.2,
                        timeout=90
                    )
                    
                    phase = progress_phases[-1]
                    status_text.markdown(f"### {phase['icon']} {T['report_complete']}")
                    sub_text.caption(T["beautifying"])
                    progress_bar.progress(1.0)
                    time.sleep(0.5)
                    
                    final_result = clean_format(response.choices[0].message.content)
                    st.session_state.current_result = final_result
                    inc_trial_used(user_id)
                    st.success(T["success"])
                    st.rerun()
                    
                except openai.APIConnectionError:
                    status_text.empty()
                    sub_text.empty()
                    progress_bar.empty()
                    st.error("🌐 网络连接失败，请检查网络后重试")
                except openai.RateLimitError:
                    status_text.empty()
                    sub_text.empty()
                    progress_bar.empty()
                    st.error("⏱️ API 调用频率超限，请等待 30 秒后重试")
                except openai.AuthenticationError:
                    status_text.empty()
                    sub_text.empty()
                    progress_bar.empty()
                    st.error("🔑 API 密钥验证失败，请联系管理员")
                except openai.APIError as e:
                    status_text.empty()
                    sub_text.empty()
                    progress_bar.empty()
                    st.error(f"❌ 服务异常：{e.type}" if hasattr(e, 'type') else "❌ 服务异常，请稍后重试")
                except Exception:
                    status_text.empty()
                    sub_text.empty()
                    progress_bar.empty()
                    st.error(T["api_error"])

with col_preview:
    st.header(T["preview_header"])
    if st.session_state.current_result:
        st.markdown(st.session_state.current_result.replace("**", "").replace("#", ""))
        st.markdown("---")
        
        user_id = st.session_state.get("user_id")
        lic = get_user_license(user_id) if user_id else None
        if lic and lic['plan_type'] != 'free':
            word_data = export_to_word(
                st.session_state.current_result, 
                product_name or "8D_Report", 
                st.session_state.lang
            )
            st.download_button(
                label=T["download_btn"],
                data=word_data,
                file_name=f"8D_Report_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.info(T["export_disabled"])
    else:
        st.info("👈 输入问题描述后点击生成" if st.session_state.lang == "zh" else "👈 Enter description and click generate")
