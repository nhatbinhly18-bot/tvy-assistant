import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# 1. 网页基础配置
st.set_page_config(page_title="体卫艺办公助手", page_icon="🚀", layout="centered")

# --- 🔙 还原回最初的经典紧凑版 CSS ---
st.markdown("""
<style>
    /* 隐藏顶部菜单和页脚 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    @media (min-width: 769px) { header {visibility: hidden;} }
    
    /* 紧凑布局设置 */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 0.75rem;
        padding-right: 0.75rem;
    }
    
    /* 移动端字体微调 */
    @media (max-width: 768px) {
        html, body, [class*="css"] { font-size: 14px !important; }
        h1 { font-size: 1.5rem !important; margin-bottom: 0.5rem !important; }
        button { font-size: 0.9rem !important; }
    }
    
    /* 经典绿色按钮 */
    div.stButton > button:first-child[kind="primary"] {
        background-color: #28a745;
        border-color: #28a745;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# --- 🔒 核心配置 ---
CONTACT_PASSWORD = "lhjy" 
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"

# 初始化状态
if "contacts_authenticated" not in st.session_state:
    st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state:
    st.session_state.parseddata_doc = None
if "step" not in st.session_state:
    st.session_state.step = 1 

# 3. 侧边栏
with st.sidebar:
    st.header("⚙️ 体卫艺办公助手")
    st.success("● AI 核心已连接")
    st.caption("维护者：孙沛 | 龙华区教育局")
    st.divider()
    mode = st.radio("功能切换：", ["📝 领导公务单自动生成器", "🔍 龙华学校查号台"])
    if st.button("🔒 退出并锁定"):
        st.session_state.contacts_authenticated = False
        st.rerun()

# ----------------- 模块一：领导公务单生成器 -----------------
if mode == "📝 领导公务单自动生成器":
    st.markdown("# 🚀 领导公务单自动生成器")
    
    if st.session_state.step == 1:
        st.info("💡 请一次性说清：时间、地点、名称、人数、对接人、领导、背景及议程。")
        user_input = st.text_area("✍️ 请输入活动描述：", height=150, placeholder="在此粘贴内容...")
        
        if st.button("✨ 立即智能填表并生成 Word", type="primary"):
            if user_input:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                with st.spinner("解析中..."):
                    try:
                        full_prompt = f"解析公文要素并以JSON返回：{user_input}。字段含：title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。"
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct",
                            messages=[{"role": "user", "content": full_prompt}],
                            response_format={'type': 'json_object'}
                        )
                        st.session_state.parseddata_doc = json.loads(chat_completion.choices[0].message.content)
                        st.session_state.step = 2
                        st.rerun()
                    except Exception as e: st.error(f"解析失败: {e}")

    elif st.session_state.step == 2:
        d = st.session_state.parseddata_doc
        # 使用原汁原味的 border=True 容器
        with st.container(border=True):
            st.markdown("### 🧐 核心要素预览与微调")
            t = st.text_input("📝 政务活动名称", d.get("title", ""))
            c = st.text_area("📄 理由背景", d.get("content", ""), height=80)
            
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list): agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            a = st.text_area("📋 议程", agenda_val or "1. 专题汇报\n2. 座谈交流\n3. 领导讲话", height=100)
            
            col1, col2 = st.columns(2)
            with col1:
                tm = st.text_input("⏰ 时间", d.get("time", ""))
                dr = st.text_input("⏳ 时长", d.get("duration", "1小时"))
            with col2:
                st.caption("时间可否调整：☑否")
                ct = st.text_input("👤 对接人", d.get("contact", "孙沛"))

            col3, col4, col5 = st.columns([2, 1, 1])
            with col3: pl = st.text_input("📍 地点", d.get("place", ""))
            with col4: nm = st.text_input("👥 人数", d.get("num", ""))
            with col5: pj = st.selectbox("📽️ 投影仪", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            dist_l = st.text_input("1. 拟请出席的区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("2. 拟请协调出席的局领导", d.get("bur_leader", ""))
            oth = st.text_input("建议参加单位", d.get("others") or "体卫艺劳科")

        col_final_back, col_final_down = st.columns([1, 2])
        with col_final_back:
            if st.button("⬅️ 返回上一步"):
                st.session_state.step = 1
                st.rerun()

        with col_final_down:
            try:
                tpl = DocxTemplate("申报单模板.docx")
                tpl.render({"title":t,"content":c,"agenda":a,"time":tm,"duration":dr,"place":pl,"num":nm,"contact":ct,"projector":pj,"dist_leader":dist_l,"bur_leader":bur_l,"others":oth})
                bio = io.BytesIO()
                tpl.save(bio)

                # --- 核心保留：1229-灵芝主任命名逻辑 ---
                mmdd = datetime.now().strftime("%m%d") 
                raw_leader = dist_l.strip() if dist_l.strip() else bur_l.strip()
                if not raw_leader:
                    leader_display = "领导"
                else:
                    first = raw_leader.split('、')[0]
                    leader_display = "灵芝主任" if ("杨灵芝" in first or "灵芝" in first) else first
                
                filename = f"{mmdd}-{leader_display}-体卫艺劳科-{t}.docx"

                st.download_button(label="💾 确认无误，导出 Word", data=bio.getvalue(), file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
            except Exception as e: st.error(f"失败: {e}")

# ----------------- 模块二：龙华学校查号台 (原样回归) -----------------
else:
    st.markdown("### 🔍 龙华学校查号台")
    if not st.session_state.contacts_authenticated:
        pwd = st.text_input("请输入授权密码", type="password")
        if st.button("验证登录", type="primary"):
            if pwd == CONTACT_PASSWORD:
                st.session_state.contacts_authenticated = True
                st.rerun()
            else: st.error("密码错误")
        st.stop()

    @st.cache_data
    def load_contacts():
        try: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
        except: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    q = st.text_input("请输入关键词搜索：", placeholder="例如：龙华中学")
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
    else:
        st.dataframe(df.head(5), use_container_width=True, hide_index=True)
