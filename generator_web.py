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
# 2. Cookies：试用次数持久化（刷新不重置）
# =========================================================
MAX_TRIAL_GENERATIONS = 3          # 试用允许生成次数（你要 3 次就保持 3）
TRIAL_COOKIE_NAME = "trial_count"  # cookie key（内部用）
COOKIE_PREFIX = "ai8d/8d-report/"  # 防止共享域下名字冲突（建议固定）

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


def get_trial_used() -> int:
    """从 cookie 读取已用次数（刷新/重开浏览器仍然存在）"""
    return _safe_int(cookies.get(TRIAL_COOKIE_NAME), 0)


def set_trial_used(val: int):
    """写入 cookie 并立即保存"""
    cookies[TRIAL_COOKIE_NAME] = str(max(0, val))
    cookies.save()


def inc_trial_used() -> int:
    """已用次数 +1，并保存，返回最新值"""
    new_val = get_trial_used() + 1
    set_trial_used(new_val)
    return new_val


def trial_remaining() -> int:
    """剩余次数"""
    return max(0, MAX_TRIAL_GENERATIONS - get_trial_used())


# =========================================================
# 3. 授权状态（仍保持会话级：trial/active）
#    说明：你这版激活逻辑依然是“长度>=8 即激活”，不做后端校验。
# =========================================================
if "license_type" not in st.session_state:
    st.session_state.license_type = "trial"  # trial / active

if "history" not in st.session_state:
    st.session_state.history = []


# =========================================================
# 4. 侧边栏：状态/激活/历史
# =========================================================
def license_panel():
    with st.sidebar:
        st.header("系统状态")

        if st.session_state.license_type == "trial":
            used = get_trial_used()
            rem = trial_remaining()

            st.warning(f"⚠️ 试用版：允许生成 {MAX_TRIAL_GENERATIONS} 次（已用 {used}，剩余 {rem}）")
            st.caption("试用版：可完整预览生成内容；正式版：支持 Word 导出 + 无限次生成。")

            if rem <= 0:
                st.error("❌ 试用次数已用完：请输入激活码继续使用。")

            code = st.text_input("输入激活码", type="password")
            if st.button("激活正式版", use_container_width=True):
                if len(code) >= 8:
                    st.session_state.license_type = "active"
                    st.success("✅ 激活成功（当前会话）")
                else:
                    st.error("激活码格式无效（至少 8 位）")

            # （可选）给你自己一个“重置试用”的隐藏入口，方便演示
            with st.expander("（仅管理员）试用次数管理", expanded=False):
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("重置为 0", use_container_width=True):
                        set_trial_used(0)
                        st.success("已重置为 0")
                        st.rerun()
                with col_b:
                    if st.button("扣回 1 次", use_container_width=True):
                        set_trial_used(max(0, get_trial_used() - 1))
                        st.success("已扣回 1 次")
                        st.rerun()

        else:
            st.success("✅ 正式版：无限次生成 + 支持 Word 导出")

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
        placeholder="请使用 5W2H 描述异常事实（越具体越好）"
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

        # =========================================================
        # ✅ 试用次数限制（关键修复）
        # 规则：trial 且 已用次数 >= MAX_TRIAL -> 拦截
        # 这样 3 次就是真 3 次，不会提前少一次
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
  A. 为什么会发生（4M1E + 仅异常项 5Why）
  B. 为什么会流出（5Why）
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

                # ✅ 生成成功后再扣次数（顺序正确）
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

        # ✅ 只有正式版允许导出 Word
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
