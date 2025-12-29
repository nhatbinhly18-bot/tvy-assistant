import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# 1. 网页基础配置
st.set_page_config(page_title="体卫艺办公助手", page_icon="🚀", layout="centered")

# --- Mobile Optimization / Custom CSS ---
st.markdown("""
<style>
    /* 隐藏顶部菜单和页脚，但保留移动端侧边栏按钮 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 桌面端隐藏header，移动端保留以便访问侧边栏 */
    @media (min-width: 769px) {
        header {visibility: hidden;}
    }
    
    /* 调整移动端内边距 */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 0.75rem;
        padding-right: 0.75rem;
    }
    
    /* 侧边栏优化 */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
    
    /* 移动端字体优化 */
    @media (max-width: 768px) {
        html, body, [class*="css"] {
            font-size: 14px !important;
        }
        h1 { font-size: 1.5rem !important; margin-bottom: 0.5rem !important; }
        h2 { font-size: 1.25rem !important; margin-bottom: 0.5rem !important; }
        h3 { font-size: 1.1rem !important; margin-bottom: 0.5rem !important; }
        [data-testid="stWidgetLabel"] { font-size: 0.85rem !important; }
        button { font-size: 0.9rem !important; padding: 0.4rem 0.8rem !important; }
        textarea, input { font-size: 0.9rem !important; }
        [data-testid="stSidebar"] { font-size: 0.85rem !important; }
        .element-container { margin-bottom: 0.5rem !important; }
        [data-testid="stExpander"] summary { font-size: 0.9rem !important; }
    }
</style>
""", unsafe_allow_html=True)

# --- 🔒 通讯录专属密码 ---
CONTACT_PASSWORD = "lhjy" 

# 2. 核心配置
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"

# 初始化状态
if "contacts_authenticated" not in st.session_state:
    st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state:
    st.session_state.parseddata_doc = None
if "step" not in st.session_state:
    st.session_state.step = 1 
if "polished_text" not in st.session_state:
    st.session_state.polished_text = None
if "original_input" not in st.session_state:
    st.session_state.original_input = ""

# 3. 侧边栏导航
with st.sidebar:
    st.header("⚙️ 体卫艺办公助手")
    st.success("● AI 核心已连接")
    
    st.info("""
    **💡 使用小提示：** 本助手集成两大核心功能：
    1. **公务单生成**：智能解析文字生成 Word。
    2. **学校查号台**：全区学校通讯录快速查询。
    """)
    st.caption("维护者：孙沛 | 龙华区教育局体卫艺专用")
    st.divider()
    mode = st.radio("功能切换：", ["📝 领导公务单自动生成器", "🔍 龙华学校查号台"])
    
    if st.button("🔒 退出并锁定"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ----------------- 模块一：领导公务单生成器 -----------------
if mode == "📝 领导公务单自动生成器":
    st.warning("👆 点击左上角 **>>** 可切换到「查号台」")
    st.markdown("# 🚀 领导公务单自动生成器")
    st.markdown("<div style='font-size: 18px; margin: 0.3rem 0; line-height: 1.4;'>欢迎使用！本工具旨在帮您一键完成体卫艺政务活动申报。</div>", unsafe_allow_html=True)

    st.info("""
    **💡 请一次性说清：** 时间、地点、会议名称、人数、对接人、领导、参加部门、背景及议程。
    **参考范例：** 明天上午10点在二楼多功能厅有个生涯教育座谈会，大概20人，孙沛对接，时长1小时，邀请灵芝主任参加
    """)

    # 绿色按钮样式
    st.markdown("""
    <style>
        div.stButton > button:first-child[kind="primary"] {
            background-color: #28a745; border-color: #28a745; color: white;
        }
    </style>
    """, unsafe_allow_html=True)

    if st.session_state.step == 1:
        user_input = st.text_area("✍️ 请输入活动描述：", height=150, placeholder="请在此粘贴或输入内容...", key="input_doc")
        
        if st.button("✨ 立即智能填表并生成 Word", type="primary"):
            if not user_input:
                st.warning("内容不能为空。")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                st.session_state.original_input = user_input
                current_date_str = datetime.now().strftime("%Y年%m月%d日")
                weekday = datetime.now().strftime("%w")
                
                with st.spinner("正在解析要素..."):
                    full_prompt = f"你现在是龙华教育局资深笔杆子。根据输入解析公文要素：{user_input}。今天是{current_date_str}。请严格按 JSON 格式返回字段：title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。"
                    try:
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct",
                            messages=[{"role": "user", "content": full_prompt}],
                            response_format={'type': 'json_object'}
                        )
                        result = json.loads(chat_completion.choices[0].message.content)
                        st.session_state.parseddata_doc = result
                        st.session_state.step = 2
                        st.rerun()
                    except Exception as e:
                        st.error(f"解析出错：{e}")

    elif st.session_state.step == 2 and st.session_state.parseddata_doc:
        d = st.session_state.parseddata_doc
        with st.container(border=True):
            st.markdown("### 🧐 核心要素预览与微调")
            t = st.text_input("📝 政务活动名称", d.get("title", ""))
            c = st.text_area("📄 政务活动申请理由、背景", d.get("content", ""), height=80)
            
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list): agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            if not agenda_val: agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("📋 议程", agenda_val, height=120)
            
            col1, col2 = st.columns(2)
            with col1:
                tm = st.text_input("⏰ 时间", d.get("time", ""))
                dr = st.text_input("⏳ 会议时长", d.get("duration", "1小时"))
            with col2:
                st.caption("时间可否调整：☑否")
                ct = st.text_input("👤 公务对接人", d.get("contact", "孙沛"))

            col3, col4, col5 = st.columns([2, 1, 1])
            with col3: pl = st.text_input("📍 地点", d.get("place", ""))
            with col4: nm = st.text_input("👥 人数", d.get("num", ""))
            with col5: pj = st.selectbox("📽️ 投影仪", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            st.markdown("---")
            dist_l = st.text_input("1. 拟请出席的区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("2. 拟请办公室协调出席的局领导", d.get("bur_leader", ""))
            oth = st.text_input("建议参加单位(部门)", d.get("others") or "体卫艺劳科")

        col_final_back, col_final_down = st.columns([1, 2])
        with col_final_back:
            if st.button("⬅️ 返回上一步"):
                st.session_state.step = 1
                st.rerun()

        with col_final_down:
            try:
                final_data = {
                    "title": t, "content": c, "agenda": a, "time": tm, 
                    "duration": dr, "place": pl, "num": nm, "contact": ct, 
                    "projector": pj, "dist_leader": dist_l, "bur_leader": bur_l, "others": oth
                }
                tpl = DocxTemplate("申报单模板.docx")
                tpl.render(final_data)
                bio = io.BytesIO()
                tpl.save(bio)

                # --- 核心修改：文件名命名逻辑 ---
                mmdd = datetime.now().strftime("%m%d") 
                # 优先级：区领导 > 局领导
                raw_leader = dist_l.strip() if dist_l.strip() else bur_l.strip()
                
                if not raw_leader:
                    leader_display = "领导"
                else:
                    # 提取名字
                    first_name = raw_leader.split('、')[0] if '、' in raw_leader else raw_leader
                    # 杨灵芝 特殊映射
                    if "杨灵芝" in first_name or "灵芝" in first_name:
                        leader_display = "灵芝主任"
                    else:
                        leader_display = first_name

                # 最终格式：1229-灵芝主任-体卫艺劳科-活动名称.docx
                filename = f"{mmdd}-{leader_display}-体卫艺劳科-{t}.docx"
                # ------------------------------

                st.download_button(
                    label="💾 确认无误，导出 Word",
                    data=bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"生成失败：{e}")

# 🔍 龙华学校查号台 (此处省略后续模块逻辑，保持原样即可)
# ----------------- 模块二：龙华学校查号台 -----------------
else:
    st.markdown("### 🔍 龙华学校查号台")
    if not st.session_state.contacts_authenticated:
        st.info("🔒 为了数据安全，访问通讯录需要授权。")
        pwd = st.text_input("请输入授权密码", type="password", help="请向管理员获取密码")
        if st.button("验证登录", type="primary"):
            if pwd == CONTACT_PASSWORD:
                st.session_state.contacts_authenticated = True
                st.rerun()
            else:
                st.error("密码错误，请重试。")
        st.stop()

    @st.cache_data
    def load_contacts():
        try:
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
        except:
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    q = st.text_input("请输入学校名或人名关键词：", placeholder="例如：龙华中学 或 张三")
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
    else:
        st.caption("👆 在上方输入关键词开始搜索")
