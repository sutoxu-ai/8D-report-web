#!/usr/bin/env python3
"""
多平台销售管理后台
用于生成、统计、管理各渠道激活码
"""

import streamlit as st
from supabase import create_client
from datetime import datetime, timedelta
import pandas as pd
import secrets

# ==================== 页面配置 ====================
st.set_page_config(page_title="8D 系统 - 多平台管理后台", page_icon="📊", layout="wide")

# ==================== 初始化 Supabase ====================
@st.cache_resource
def init_supabase():
    """初始化 Supabase 客户端"""
    try:
        supabase_url = st.secrets["SUPABASE_URL"]
        supabase_key = st.secrets["SUPABASE_SERVICE_ROLE"]
        return create_client(supabase_url, supabase_key)
    except Exception as e:
        st.error(f"Supabase 连接失败: {e}")
        return None

supabase = init_supabase()

# ==================== 会话状态初始化 ====================
if "admin_logged" not in st.session_state:
    st.session_state.admin_logged = False

# ==================== 侧边栏 - 管理员登录 ====================
with st.sidebar:
    st.title("🔐 管理员登录")
    
    if not st.session_state.admin_logged:
        admin_pwd = st.text_input("管理密码", type="password")
        
        if st.button("登录", use_container_width=True):
            correct_pwd = st.secrets.get("ADMIN_PASSWORD", "admin888")
            if admin_pwd == correct_pwd:
                st.session_state.admin_logged = True
                st.rerun()
            else:
                st.error("密码错误")
        st.stop()
    
    st.success("✅ 已登录")
    
    if st.button("🚪 退出登录", use_container_width=True):
        st.session_state.admin_logged = False
        st.rerun()
    
    st.markdown("---")
    
    menu = st.radio(
        "📋 导航菜单",
        ["📊 数据统计", "➕ 生成激活码", "📦 渠道管理", "🔍 查询激活码"]
    )

# ==================== 工具函数 ====================
def get_all_codes():
    """获取所有激活码"""
    if supabase is None:
        return []
    try:
        r = supabase.table("activation_codes").select("*").order("created_at", desc=True).execute()
        return r.data if r.data else []
    except Exception as e:
        st.error(f"获取数据失败: {e}")
        return []

def get_channel_name(channel_code):
    """获取渠道中文名称"""
    channel_map = {
        'taobao': '🛒 淘宝',
        'xiaohongshu': '📕 小红书',
        'wechat': '💬 微信',
        'xianyu': '🐟 闲鱼',
        'douyin': '🎵 抖音',
        'bilibili': '📺 B站',
        'other': '📦 其他',
        'unknown': '❓ 未知'
    }
    return channel_map.get(channel_code, channel_code)

def get_plan_name(plan_code):
    """获取版本中文名称"""
    plan_map = {
        'trial': '试用版（7天）',
        'pro': '专业版（1年）',
        'enterprise': '企业版（永久）'
    }
    return plan_map.get(plan_code, plan_code)

# ==================== 页面 1：数据统计 ====================
if menu == "📊 数据统计":
    st.title("📊 多平台销售数据统计")
    
    codes = get_all_codes()
    
    if not codes:
        st.info("暂无数据，请先生成激活码")
        st.stop()
    
    # 总体指标
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total = len(codes)
    used = sum(1 for c in codes if c.get('is_used', False))
    unused = total - used
    
    # 计算各版本收入（估算）
    price_map = {'trial': 0, 'pro': 299, 'enterprise': 999}
    revenue = 0
    for c in codes:
        if c.get('is_used', False):
            revenue += price_map.get(c.get('plan_type', 'pro'), 299)
    
    with col1:
        st.metric("📦 总激活码", total)
    with col2:
        st.metric("✅ 已使用", used)
    with col3:
        st.metric("⭕ 未使用", unused)
    with col4:
        st.metric("📈 使用率", f"{used/total*100:.1f}%" if total > 0 else "0%")
    with col5:
        st.metric("💰 估算收入", f"¥{revenue:,}")
    
    st.markdown("---")
    
    # 按渠道统计
    st.subheader("📈 各渠道销售情况")
    
    channel_stats = {}
    for c in codes:
        ch = c.get('channel', 'unknown')
        if ch not in channel_stats:
            channel_stats[ch] = {'total': 0, 'used': 0, 'revenue': 0}
        channel_stats[ch]['total'] += 1
        if c.get('is_used', False):
            channel_stats[ch]['used'] += 1
            channel_stats[ch]['revenue'] += price_map.get(c.get('plan_type', 'pro'), 299)
    
    # 构建表格数据
    table_data = []
    for ch, stats in channel_stats.items():
        table_data.append({
            '渠道': get_channel_name(ch),
            '总数量': stats['total'],
            '已使用': stats['used'],
            '未使用': stats['total'] - stats['used'],
            '使用率': f"{stats['used']/stats['total']*100:.1f}%" if stats['total'] > 0 else "0%",
            '估算收入': f"¥{stats['revenue']:,}"
        })
    
    df = pd.DataFrame(table_data)
    st.dataframe(df, use_container_width=True, hide_index=True)
    
    # 按版本统计
    st.markdown("---")
    st.subheader("📊 按版本统计")
    
    plan_stats = {}
    for c in codes:
        plan = c.get('plan_type', 'unknown')
        if plan not in plan_stats:
            plan_stats[plan] = {'total': 0, 'used': 0}
        plan_stats[plan]['total'] += 1
        if c.get('is_used', False):
            plan_stats[plan]['used'] += 1
    
    plan_data = []
    for plan, stats in plan_stats.items():
        plan_data.append({
            '版本': get_plan_name(plan),
            '总数量': stats['total'],
            '已使用': stats['used'],
            '使用率': f"{stats['used']/stats['total']*100:.1f}%" if stats['total'] > 0 else "0%"
        })
    
    df_plan = pd.DataFrame(plan_data)
    st.dataframe(df_plan, use_container_width=True, hide_index=True)

# ==================== 页面 2：生成激活码 ====================
elif menu == "➕ 生成激活码":
    st.title("➕ 批量生成激活码")
    
    if supabase is None:
        st.error("数据库连接失败，请检查配置")
        st.stop()
    
    col1, col2 = st.columns(2)
    
    with col1:
        channel = st.selectbox(
            "📌 销售渠道",
            ["taobao", "xiaohongshu", "wechat", "xianyu", "douyin", "bilibili", "other"],
            format_func=get_channel_name
        )
        
        plan_type = st.selectbox(
            "📦 版本类型",
            ["trial", "pro", "enterprise"],
            format_func=get_plan_name
        )
        
        duration_map = {
            'trial': 7,
            'pro': 365,
            'enterprise': 9999
        }
        
        count = st.number_input(
            "🔢 生成数量",
            min_value=1,
            max_value=1000,
            value=10
        )
        
        expire_date = st.date_input(
            "📅 激活码有效期（过期未使用则失效）",
            datetime.now() + timedelta(days=365)
        )
        
        batch_note = st.text_input(
            "📝 批次备注（可选）",
            placeholder="例如：双11活动批次"
        )
    
    with col2:
        st.info(f"""
        **📋 生成预览**
        
        | 项目 | 值 |
        |------|-----|
        | 渠道 | {get_channel_name(channel)} |
        | 版本 | {get_plan_name(plan_type)} |
        | 数量 | {count} 个 |
        | 单码有效期 | {duration_map[plan_type]} 天 |
        | 激活码过期 | {expire_date} |
        | 批次备注 | {batch_note or '无'} |
        """)
    
    st.markdown("---")
    
    if st.button("🚀 生成激活码", type="primary", use_container_width=True):
        chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'
        generated = []
        failed = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        for i in range(count):
            status_text.text(f"正在生成第 {i+1}/{count} 个激活码...")
            
            # 生成激活码
            prefix = {'trial': '8D30', 'pro': '8D8P', 'enterprise': '8D8E'}[plan_type]
            part1 = ''.join(secrets.choice(chars) for _ in range(4))
            part2 = ''.join(secrets.choice(chars) for _ in range(4))
            body = f"{prefix}-{part1}-{part2}"
            check = str(sum(ord(c) * (idx + 1) for idx, c in enumerate(body)) % 10)
            code = f"{body}-{check}"
            
            data = {
                "code": code,
                "plan_type": plan_type,
                "duration_days": duration_map[plan_type],
                "channel": channel,
                "expire_date": expire_date.isoformat(),
                "batch_id": batch_id,
                "batch_note": batch_note if batch_note else None,
                "is_used": False,
                "created_at": datetime.now().isoformat()
            }
            
            try:
                supabase.table("activation_codes").insert(data).execute()
                generated.append(code)
            except Exception as e:
                failed += 1
                st.warning(f"生成 {code} 失败: {e}")
            
            progress_bar.progress((i + 1) / count)
        
        status_text.empty()
        
        if generated:
            st.success(f"✅ 成功生成 {len(generated)} 个激活码" + (f"，失败 {failed} 个" if failed > 0 else ""))
            
            # 导出 CSV
            csv_content = "激活码,版本,渠道,有效期,批次号\n"
            for code in generated:
                csv_content += f"{code},{plan_type},{channel},{expire_date},{batch_id}\n"
            
            st.download_button(
                label=f"📥 导出 CSV ({len(generated)}个激活码)",
                data=csv_content,
                file_name=f"activation_codes_{channel}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # 显示生成的激活码
            with st.expander("🔍 查看生成的激活码"):
                st.code("\n".join(generated))
        else:
            st.error("生成失败，请检查数据库连接")

# ==================== 页面 3：渠道管理 ====================
elif menu == "📦 渠道管理":
    st.title("📦 渠道库存管理")
    
    codes = get_all_codes()
    
    if not codes:
        st.info("暂无数据")
        st.stop()
    
    # 按渠道统计
    channel_inventory = {}
    for c in codes:
        ch = c.get('channel', 'unknown')
        if ch not in channel_inventory:
            channel_inventory[ch] = {
                'total': 0,
                'used': 0,
                'unused': 0,
                'trial': 0,
                'pro': 0,
                'enterprise': 0
            }
        channel_inventory[ch]['total'] += 1
        if c.get('is_used', False):
            channel_inventory[ch]['used'] += 1
        else:
            channel_inventory[ch]['unused'] += 1
        
        plan = c.get('plan_type', 'unknown')
        if plan in ['trial', 'pro', 'enterprise']:
            channel_inventory[ch][plan] += 1
    
    # 库存总览表格
    inventory_data = []
    for ch, stats in channel_inventory.items():
        inventory_data.append({
            '渠道': get_channel_name(ch),
            '总库存': stats['total'],
            '已使用': stats['used'],
            '剩余': stats['unused'],
            '试用版': stats['trial'],
            '专业版': stats['pro'],
            '企业版': stats['enterprise'],
            '使用率': f"{stats['used']/stats['total']*100:.1f}%" if stats['total'] > 0 else "0%"
        })
    
    df = pd.DataFrame(inventory_data)
    st.dataframe(df, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # 库存预警
    st.subheader("⚠️ 库存预警")
    
    low_stock_threshold = st.slider("低库存阈值", min_value=5, max_value=50, value=20)
    
    low_stock = [
        (ch, stats['unused']) 
        for ch, stats in channel_inventory.items() 
        if stats['unused'] < low_stock_threshold
    ]
    
    if low_stock:
        for ch, unused in low_stock:
            st.warning(f"⚠️ {get_channel_name(ch)} 渠道剩余库存仅 {unused} 个，请及时补充！")
    else:
        st.success(f"✅ 所有渠道库存充足（均大于 {low_stock_threshold} 个）")
    
    st.markdown("---")
    
    # 快速生成补充库存
    st.subheader("⚡ 快速补货")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        quick_channel = st.selectbox(
            "选择渠道",
            list(channel_inventory.keys()),
            format_func=get_channel_name,
            key="quick_channel"
        )
    with col2:
        quick_plan = st.selectbox(
            "选择版本",
            ["pro", "trial", "enterprise"],
            format_func=get_plan_name,
            key="quick_plan"
        )
    with col3:
        quick_count = st.number_input("数量", min_value=1, max_value=100, value=10, key="quick_count")
    with col4:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 快速生成", use_container_width=True):
            st.switch_page("pages/生成激活码")  # 跳转到生成页面

# ==================== 页面 4：查询激活码 ====================
elif menu == "🔍 查询激活码":
    st.title("🔍 激活码查询")
    
    col1, col2 = st.columns(2)
    
    with col1:
        search_code = st.text_input("🔑 输入激活码查询", placeholder="例如：8D8P-XXXX-XXXX-X")
    with col2:
        search_batch = st.text_input("📦 输入批次号查询", placeholder="例如：20240115_143022")
    
    if search_code or search_batch:
        q = supabase.table("activation_codes").select("*")
        
        if search_code:
            q = q.eq("code", search_code.strip().upper())
        elif search_batch:
            q = q.eq("batch_id", search_batch.strip())
        
        r = q.execute()
        
        if r.data:
            if len(r.data) == 1:
                code = r.data[0]
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    status = "✅ 已使用" if code.get('is_used') else "⭕ 未使用"
                    st.metric("状态", status)
                with col2:
                    st.metric("版本", get_plan_name(code.get('plan_type', 'unknown')))
                with col3:
                    st.metric("渠道", get_channel_name(code.get('channel', 'unknown')))
                
                col4, col5, col6 = st.columns(3)
                
                with col4:
                    st.metric("有效期", f"{code.get('duration_days', 0)} 天")
                with col5:
                    if code.get('is_used'):
                        used_at = code.get('used_at', '')
                        st.metric("使用时间", used_at[:19] if used_at else '未知')
                    else:
                        st.metric("过期日期", code.get('expire_date', '无'))
                with col6:
                    st.metric("使用者", code.get('used_by', '—'))
                
                st.markdown("---")
                st.subheader("📋 完整信息")
                st.json(code)
                
            else:
                st.success(f"找到 {len(r.data)} 个激活码（批次：{search_batch}）")
                
                # 显示列表
                batch_codes = []
                for c in r.data:
                    batch_codes.append({
                        '激活码': c['code'],
                        '状态': '✅ 已使用' if c.get('is_used') else '⭕ 未使用',
                        '使用者': c.get('used_by', '—'),
                        '使用时间': c.get('used_at', '')[:19] if c.get('used_at') else '—'
                    })
                
                df_batch = pd.DataFrame(batch_codes)
                st.dataframe(df_batch, use_container_width=True, hide_index=True)
                
                # 导出批次
                csv = "激活码,状态,使用者,使用时间\n"
                for c in r.data:
                    csv += f"{c['code']},{c.get('is_used', False)},{c.get('used_by', '')},{c.get('used_at', '')}\n"
                
                st.download_button(
                    f"📥 导出批次 ({len(r.data)}个)",
                    csv,
                    f"batch_{search_batch}.csv",
                    use_container_width=True
                )
        else:
            st.error("未找到相关激活码")

# ==================== 页脚 ====================
st.markdown("---")
st.caption(f"📊 8D 报告系统 - 多平台管理后台 | 最后更新: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
