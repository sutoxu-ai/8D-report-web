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
        # 如果不存在，尝试创建免费许可证
        return create_free_license(user_id)
    except Exception:
        return None

def clear_license_cache(user_id):
    """清除特定用户的缓存"""
    get_cached_license.clear()

# ==================== 页面配置 ====================
st.set_page_config(page_title="8D 报告智能生成助手", page_icon="📊", layout="wide")

# ==================== 多语言文本 ====================
TEXT = {
    "zh": {
        "lang_label": "Language", "lang_zh": "中文", "lang_en": "English",
        "system_status": "系统状态", "pro_version": "✅ 正式版", "trial_version": "⚠️ 试用版",
        "license_valid_until": "📅 有效期至 {exp}", "trial_used": "📊 已使用 {used} 次 / 共 {total} 次",
        "trial_exhausted": "❌ 试用次数已用完", "activate_title": "🔑 授权 / 续费",
        "activate_code_hint": "激活码", "activate_btn": "立即激活",
        "activate_success": "✅ 激活成功，有效期一年", "activate_fail": "❌ 激活码无效",
        "login_required": "🔒 请先登录", "logout": "退出登录",
        "main_title": "📊 8D 报告智能生成助手", "input_header": "📝 输入基本信息",
        "product_name": "产品型号 / 名称", "customer": "客户名称",
        "problem_desc": "不良现象描述",
        "problem_placeholder": "请使用 5W2H 方法描述问题",
        "occur_date": "发现日期", "defect_qty": "不良数量", "severity": "严重程度",
        "severity_low": "低", "severity_medium": "中", "severity_high": "高", "severity_critical": "危急",
        "industry_std": "适用标准", "team_members": "团队成员（可选）",
        "team_placeholder": "例：张明 (组长), 李华 (工程)",
        "generate_btn": "🚀 生成 8D 报告", "generating": "8D 报告正在生成中...",
        "preview_header": "📄 报告预览", "download_btn": "📥 导出 Word 报告",
        "export_disabled": "🔒 激活正式版后可导出 Word",
        "no_desc": "❌ 请输入不良现象描述",
        "trial_exhausted_error": "❌ 试用次数已用完", "api_error": "❌ 服务异常",
        "success": "✅ 报告生成完成！", "word_title": "8D 问题纠正与预防措施报告"
    },
    "en": {
        "lang_label": "Language", "lang_zh": "中文", "lang_en": "English",
        "system_status": "System Status", "pro_version": "✅ Pro Version", "trial_version": "⚠️ Trial Version",
        "license_valid_until": "📅 Valid until {exp}", "trial_used": "📊 Used {used} / {total}",
        "trial_exhausted": "❌ Trial exhausted", "activate_title": "🔑 License / Renew",
        "activate_code_hint": "Activation Code", "activate_btn": "Activate",
        "activate_success": "✅ Activated successfully", "activate_fail": "❌ Invalid code",
        "login_required": "🔒 Please login", "logout": "Logout",
        "main_title": "📊 8D Report Generator", "input_header": "📝 Input Information",
        "product_name": "Product Name / Model", "customer": "Customer Name",
        "problem_desc": "Problem Description",
        "problem_placeholder": "Please use 5W2H method",
        "occur_date": "Occurrence Date", "defect_qty": "Defect Quantity", "severity": "Severity",
        "severity_low": "Low", "severity_medium": "Medium", "severity_high": "High", "severity_critical": "Critical",
        "industry_std": "Standard", "team_members": "Team Members (Optional)",
        "team_placeholder": "e.g., Zhang(Leader), Li(Eng)",
        "generate_btn": "🚀 Generate 8D Report", "generating": "Generating report...",
        "preview_header": "📄 Report Preview", "download_btn": "📥 Export Word",
        "export_disabled": "🔒 Activate to export",
        "no_desc": "❌ Please enter description",
        "trial_exhausted_error": "❌ Trial exhausted", "api_error": "❌ Service error",
        "success": "✅ Report generated!", "word_title": "8D Corrective Action Report"
    }
}

SYSTEM_PROMPT = {
    "zh": "你是一位拥有 20 年经验的汽车电子行业高级质量工程师，精通 IATF 16949 标准和 8D 问题解决方法。请根据用户输入撰写专业、逻辑严密的 8D 报告。要求：1.语气专业客观 2.4M1E 分析包含人机料法环 3. 包含 5-Why 分析 4. 措施使用 [责任人 | 时间 | 状态] 格式 5. 不使用 Markdown 标记。直接输出 D1-D8 报告正文。",
    "en": "You are a Senior Quality Engineer with 20 years experience in automotive electronics, proficient in IATF 16949 and 8D methodology. Please write a professional 8D report based on user input. Requirements: 1.Professional tone 2.4M1E analysis 3.5-Why analysis 4.Use [Owner|Date|Status] format 5.No Markdown. Output D1-D8 directly."
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
    """获取用户许可证（使用缓存）"""
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
    
    # 添加 trial 类型的支持
    if lic['plan_type'] in ['free', 'trial']:  # ← 改这里
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
    """增加试用次数"""
    if not supabase:
        return
    try:
        lic = get_cached_license(user_id)
        if lic:
            new_count = lic.get('trial_used', 0) + 1
            supabase.table("licenses").update({"trial_used": new_count}).eq("user_id", user_id).execute()
            
            # 记录日志
            supabase.table("usage_logs").insert({
                "user_id": user_id,
                "action": "generate_report",
                "created_at": datetime.now().isoformat()
            }).execute()
            
            # 清除缓存
            clear_license_cache(user_id)
    except Exception as e:
        logging.error(f"更新试用次数失败：{e}")

def activate_license_code(user_id, code):
    """验证并使用激活码"""
    if not supabase:
        return False, "系统错误"
    try:
        # 1. 查询激活码
        r = supabase.table("activation_codes").select("*").eq("code", code.strip().upper()).execute()
        if not r.data:
            return False, "无效的激活码"
        
        ac = r.data[0]
        
        # 2. 检查是否已使用
        if ac.get('is_used'):
            return False, "激活码已被使用"
        
        # 3. 检查激活码是否过期
        if ac.get('expire_date'):
            if datetime.now().date() > datetime.fromisoformat(ac['expire_date']).date():
                return False, "激活码已过期"
        
        # 4. 计算有效期
        duration = ac.get('duration_days', 365)
        exp_date = (datetime.now() + timedelta(days=duration)).isoformat()
        
        # 5. 更新用户 license
        supabase.table("licenses").update({
            "plan_type": ac.get('plan_type', 'pro'),
            "license_expire": exp_date
        }).eq("user_id", user_id).execute()
        
        # 6. 标记激活码已使用
        supabase.table("activation_codes").update({
            "is_used": True,
            "used_by": user_id,
            "used_at": datetime.now().isoformat()
        }).eq("code", code.strip().upper()).execute()
        
        # 7. 清除缓存
        clear_license_cache(user_id)
        
        return True, f"激活成功！有效期至 {exp_date[:10]}"
        
    except Exception as e:
        logging.error(f"激活失败：{e}")
        return False, f"激活失败：{str(e)}"

def clean_format(text):
    if not text:
        return ""
    text = text.replace("**", "").replace("#", "")
    for i in range(1, 9):
        text = re.sub(rf'(D{i}[:：])\s*\n+\s*', rf'\1 ', text)
    text = re.sub(r'([^，,]+[（(][^）)]+[）)])\s*[，,]\s*', r'\1\n', text)
    for kw in ["What:", "Where:", "When:", "Who:", "Why:", "How many:", "How:"]:
        text = re.sub(rf'([^\n])({kw})', r'\1\n\2', text)
        text = re.sub(rf'({kw})([^\n])', r'\1\n\2', text)
    for kw in ["人：", "机：", "料：", "法：", "环："]:
        text = re.sub(rf'([^\n])({kw})', r'\1\n\n\2', text)
        text = re.sub(rf'({kw})([^\n])', r'\1\n\2', text)
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

# 语言切换回调
def on_lang_change():
    new_lang = "zh" if st.session_state.lang_radio == "中文" else "en"
    if new_lang != st.session_state.lang:
        st.session_state.lang = new_lang
        st.rerun()

T = TEXT[st.session_state.lang]

# ==================== 侧边栏 ====================
with st.sidebar:
    st.radio(
        T["lang_label"], 
        ["中文", "English"],
        index=0 if st.session_state.lang == "zh" else 1,
        horizontal=True, 
        key="lang_radio",
        on_change=on_lang_change
    )
    
    st.markdown("---")
    st.header("👤 用户登录")
    user_input = st.text_input("用户名/邮箱", key="user_input")
    
    if user_input:          # 第 304 行
        st.session_state.user_id = user_input
        st.success(f"欢迎，{user_input}")
        
        lic = get_user_license(user_input)
        if lic:
            if lic['plan_type'] == 'free':
                st.warning(f"⚠️ 试用版：剩余 {lic['trial_limit'] - lic['trial_used']} 次")
            else:
                st.success("✅ 专业版")
                if lic.get('license_expire'):
                    st.caption(f"有效期至：{datetime.fromisoformat(lic['license_expire']).strftime('%Y-%m-%d')}")
        
        if st.button(T["logout"]):
            st.session_state.user_id = None
            st.rerun()
        
        # 激活码输入 - 直接在这里，不要再判断 if user_input
        st.markdown("---")
        st.subheader(T["activate_title"])
        activate_code = st.text_input(T["activate_code_hint"], type="password", key="act_code")
        if st.button(T["activate_btn"], use_container_width=True):
            if not activate_code:
                st.error("请输入激活码")
            elif len(activate_code) < 6:
                st.error("激活码格式错误")
            else:
                success, msg = activate_license_code(user_input, activate_code)
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
    else:
        st.info(T["login_required"])

# ==================== 主页面 ====================
st.title(T["main_title"])
st.markdown("---")

col_input, col_preview = st.columns([1, 1.2])

with col_input:
    st.header(T["input_header"])
    product_name = st.text_input(T["product_name"], placeholder="例：PCB-A123")
    customer = st.text_input(T["customer"], placeholder="例：比亚迪汽车")
    problem_desc = st.text_area(T["problem_desc"], height=150, placeholder=T["problem_placeholder"])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        occur_date = st.date_input(T["occur_date"], datetime.now())
    with col2:
        defect_qty = st.number_input(T["defect_qty"], min_value=1, value=1)
    with col3:
        severity = st.selectbox(T["severity"], [T["severity_low"], T["severity_medium"], T["severity_high"], T["severity_critical"]])
    
    col4, col5 = st.columns(2)
    with col4:
        industry_std = st.selectbox(T["industry_std"], ["ISO 9001", "IATF 16949", "ISO 13485", "AS9100"], index=1)
    with col5:
        team_members = st.text_input(T["team_members"], placeholder=T["team_placeholder"])
    
    st.markdown("---")
    
    if st.button(T["generate_btn"], type="primary", use_container_width=True):
        if not st.session_state.get("user_id"):
            st.error("请先登录")
            st.stop()
        
        user_id = st.session_state.user_id
        
        if not can_generate_report(user_id):
            lic = get_user_license(user_id)
            if lic and lic['plan_type'] == 'free':
                st.error(T["trial_exhausted_error"])
            else:
                st.error("❌ 授权已过期")
            st.stop()
        
        if not problem_desc:
            st.error(T["no_desc"])
        else:
            with st.spinner(T["generating"]):
                try:
                    client = openai.OpenAI(api_key=API_KEY, base_url=BASE_URL)
                    user_prompt = (
                        f"请根据以下信息生成 8D 报告：产品：{product_name or '未提供'}, "
                        f"客户：{customer or '未提供'}, 日期：{occur_date}, 数量：{defect_qty}, "
                        f"严重程度：{severity}, 标准：{industry_std}, 团队：{team_members or '未提供'}\n\n"
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
                    final_result = clean_format(response.choices[0].message.content)
                    st.session_state.current_result = final_result
                    inc_trial_used(user_id)
                    st.success(T["success"])
                    st.rerun()
                    
                except openai.APIConnectionError:
                    st.error("🌐 网络连接失败，请检查网络后重试")
                except openai.RateLimitError:
                    st.error("⏱️ API 调用频率超限，请等待 30 秒后重试")
                except openai.AuthenticationError:
                    st.error("🔑 API 密钥验证失败，请联系管理员")
                except openai.APIError as e:
                    st.error(f"❌ 服务异常：{e.type}" if hasattr(e, 'type') else "❌ 服务异常，请稍后重试")
                except Exception:
                    st.error(T["api_error"])

with col_preview:
    st.header(T["preview_header"])
    if st.session_state.current_result:
        st.markdown(st.session_state.current_result.replace("**", "").replace("#", ""))
        st.markdown("---")
        
        user_id = st.session_state.get("user_id")
        lic = get_user_license(user_id) if user_id else None
        
        if lic and lic['plan_type'] != 'free':
            word_data = export_to_word(st.session_state.current_result, product_name or "8D_Report", st.session_state.lang)
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
