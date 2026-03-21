import streamlit as st
from io import BytesIO
from datetime import datetime, timedelta
import re

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

import openai

# ✅ Cookie 持久化（刷新不重置）
from streamlit_cookies_manager import EncryptedCookieManager


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
# 2. Cookies：试用次数持久化（刷新不重置）+ 年费到期
# =========================================================
MAX_TRIAL_GENERATIONS = 3                # 试用允许生成次数
TRIAL_COOKIE_NAME = "trial_count"        # cookie key（试用已用次数）
LICENSE_EXPIRE_COOKIE = "license_expire" # cookie key（正式版到期日：YYYY-MM-DD）
COOKIE_PREFIX = "ai8d/8d-report/"        # 防共享域冲突（建议固定）

# 年费期限（你要一年就保持365）
LICENSE_DAYS = 365

cookies = EncryptedCookieManager(
    prefix=COOKIE_PREFIX,
    password=st.secrets.get("COOKIES_PASSWORD", "CHANGE_ME_PLEASE_Use_Secrets"),
)

# 按组件要求：未 ready 前必须 stop，否则取不到 cookie 值
if not cookies.ready():
    st.stop()


def _safe_int(x, default=0):
    try:
        return int(x)
    except Exception:
        return default


def _safe_date_ymd(s: str):
    """解析 YYYY-MM-DD，失败返回 None"""
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None


def get_trial_used() -> int:
    """从 cookie 读取试用已用次数（刷新/重开浏览器仍然存在）"""
    return _safe_int(cookies.get(TRIAL_COOKIE_NAME), 0)


def set_trial_used(val: int):
    """写入试用次数 cookie 并立即保存"""
    cookies[TRIAL_COOKIE_NAME] = str(max(0, val))
    cookies.save()


def inc_trial_used() -> int:
    """试用已用次数 +1，并保存，返回最新值"""
    new_val = get_trial_used() + 1
    set_trial_used(new_val)
    return new_val


def trial_remaining() -> int:
    """试用剩余次数"""
    return max(0, MAX_TRIAL_GENERATIONS - get_trial_used())


def get_license_expire_date():
    """读取正式版到期日（date or None）"""
    expire_str = cookies.get(LICENSE_EXPIRE_COOKIE)
    if not expire_str:
        return None
    return _safe_date_ymd(expire_str)


def set_license_expire_date(d):
    """写入正式版到期日（YYYY-MM-DD）并保存"""
    cookies[LICENSE_EXPIRE_COOKIE] = d.strftime("%Y-%m-%d")
    cookies.save()


def is_license_valid_today() -> bool:
    """判断正式版是否在有效期内"""
    exp = get_license_expire_date()
    if not exp:
        return False
    return datetime.today().date() <= exp


def calc_new_expire_date_for_activation() -> str:
    """
    续费策略：
    - 若当前已有未过期到期日：从当前到期日往后 + LICENSE_DAYS
    - 若已过期或无到期日：从今天起 + LICENSE_DAYS
    """
    today = datetime.today().date()
    current_exp = get_license_expire_date()
    if current_exp and current_exp >= today:
        base = current_exp
    else:
        base = today
    new_exp = base + timedelta(days=LICENSE_DAYS)
    return new_exp.strftime("%Y-%m-%d")


# =========================================================
# 3. 授权状态（会话级：trial/active）
#    说明：是否“真正有效”由 cookie 中的到期日决定
# =========================================================
if "license_type" not in st.session_state:
    st.session_state.license_type = "trial"  # trial / active

if "history" not in st.session_state:
    st.session_state.history = []


def sync_license_state_from_cookie():
    """
    将会话状态与 cookie 到期日对齐：
    - 如果 cookie 显示有效：会话置 active
    - 如果 cookie 显示无效/过期：会话置 trial
    """
    if is_license_valid_today():
        st.session_state.license_type = "active"
    else:
        # 过期/无授权 -> 回到试用
        st.session_state.license_type = "trial"


# 每次脚本运行都同步一次（避免刷新/跨页面状态混乱）
sync_license_state_from_cookie()


# =========================================================
# 4. 侧边栏：状态/激活/历史
# =========================================================
def license_panel():
    with st.sidebar:
        st.header("系统状态")

        # 显示到期信息（如果有）
        exp = get_license_expire_date()
        if exp:
            exp_str = exp.strftime("%Y-%m-%d")
        else:
            exp_str = ""

        if st.session_state.license_type == "active":
            st.success("✅ 正式版：无限次生成 + 支持 Word 导出")
            if exp_str:
                st.info(f"📅 授权有效期至：{exp_str}")
        else:
            used = get_trial_used()
            rem = trial_remaining()

            st.warning(f"⚠️ 试用版：允许生成 {MAX_TRIAL_GENERATIONS} 次（已用 {used}，剩余 {rem}）")
            st.caption("试用版：可完整预览生成内容；正式版：支持 Word 导出 + 无限次生成。")
            if exp_str:
                # 有到期记录但已过期
                st.error(f"❌ 授权已到期（到期日：{exp_str}），已退回试用模式。")

            if rem <= 0:
                st.error("❌ 试用次数已用完：请输入激活码继续使用。")

        # 激活/续费入口：两种模式都可输入（trial 用于开通；active 用于续费延长）
        st.markdown("---")
        st.subheader("🔑 授权 / 续费")

        code = st.text_input("输入激活码", type="password", placeholder="请输入激活码")
        if st.button("立即激活 / 续费", use_container_width=True):
            # 你现在的激活码校验仍是轻量方式：长度>=8即通过
            # 如需“一码一客户”或“码池到期管理”，后续可升级为服务端校验
            if len(code) >= 8:
                new_exp_str = calc_new_expire_date_for_activation()
                new_exp_date = _safe_date_ymd(new_exp_str)

                if new_exp_date:
                    set_license_expire_date(new_exp_date)
                    sync_license_state_from_cookie()
                    st.success(f"✅ 操作成功！有效期已更新至 {new_exp_str}")
                    st.rerun()
                else:
                    st.error("到期日期计算失败，请联系管理员。")
            else:
                st.error("激活码格式无效（至少 8 位）")

        st.markdown("---")
        st.header("历史记录（最近 5 条）")
        for i, item in enumerate(st.session_state.history[-5:]):
            if st.button(f"📄 {item['title']}", key=f"h_{i}", use_container_width=True):
                st.session_state.current_result = item["content"]


# =========================================================
# 5. Word 写入
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
# 6. 主界面
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
        placeholder="推荐使用 5W2H 描述不良信息（越具体越好）"
    )

    p_name = st.text_input("产品型号 / 名称")
    cust = st.text_input("客户名称")

    r1, r2 = st.columns(2)
    o_date = r1.date_input("发现日期", datetime.now())
    qty = r2.number_input("不良数量", min_value=1)

    # ✅ 生成按钮
    if st.button("🚀 自动生成 8D 报告", use_container_width=True):

        if not p_desc:
            st.error("请先输入异常描述。")
            st.stop()

        # 每次关键操作前再次同步授权（防止到期后仍停留active）
        sync_license_state_from_cookie()

        # =========================================================
        # ✅ 试用次数限制（trial 且 已用>=MAX -> 拦截）
        # =========================================================
        if st.session_state.license_type == "trial":
            used_now = get_trial_used()
            if used_now >= MAX_TRIAL_GENERATIONS:
                st.error("❌ 试用生成次数已用完，请在左侧输入激活码继续使用。")
                st.stop()

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
  A. 为什么发生（4M1E + 仅异常项 5Why）
  B. 为什么流出（5Why）
D5. 永久纠正措施选择
D6. 实施纠正措施
D7. 防止再发生（标准化 / FMEA / 控制计划）
D8. 团队表彰

A 与 B 之间空两行。
"""

            try:
                resp = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": sys_prompt},
                        {"role": "user", "content": f"异常事实：{p_desc}"}
                    ],
                    temperature=0.2
                )

                result = resp.choices[0].message.content

                # ✅ 写入当前结果（完整展示，不截断）
                st.session_state.current_result = result

                # ✅ 写入历史
                st.session_state.history.append({
                    "title": f"{(p_name or '8D')[:10]}_{o_date}",
                    "content": result
                })

                # ✅ 生成成功后再扣试用次数（顺序正确）
                if st.session_state.license_type == "trial":
                    new_used = inc_trial_used()
                    rem = trial_remaining()
                    st.info(f"ℹ️ 本次已计入试用次数：已用 {new_used} / {MAX_TRIAL_GENERATIONS}，剩余 {rem}")

            except Exception as e:
                st.error(f"分析引擎响应失败：{str(e)}")


# ---------------- 预览 / 导出区 ----------------
with col2:
    st.header("📄 8D 报告预览")

    if "current_result" in st.session_state:
        content = st.session_state.current_result
        st.markdown(content)

        # 每次导出前也校验一次是否到期
        sync_license_state_from_cookie()

        # ✅ 只有正式版允许导出 Word（且必须未到期）
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
                file_name=f"8D_Report_{p_name or 'NA'}.docx",
                use_container_width=True
            )
        else:
            st.info("🔒 试用版可完整预览；激活后可导出 Word，并无限次生成。")

    else:
        if st.session_state.license_type == "trial":
            st.caption(f"提示：试用剩余 {trial_remaining()} / {MAX_TRIAL_GENERATIONS} 次（刷新不会重置）")
        else:
            exp = get_license_expire_date()
            if exp:
                st.caption(f"提示：当前授权有效期至 {exp.strftime('%Y-%m-%d')}")
