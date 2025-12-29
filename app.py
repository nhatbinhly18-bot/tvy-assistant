import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# 1. 网页基础配置
st.set_page_config(page_title="体卫艺办公助手", page_icon="🚀", layout="centered")

# --- 🎨 UI 深度美颜版 CSS ---
st.markdown("""
<style>
    /* 1. 整体背景与字体：换成更清爽的 App 质感背景 */
    .stApp {
        background-color: #F8F9FB;
    }
    
    /* 隐藏多余的顶部和底部元素 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* 2. 卡片式容器：让每一个功能块都像一张精美的卡片 */
    [data-testid="stVerticalBlock"] > div:has(div.stMarkdown) {
        background-color: white !important;
        padding: 24px !important;
        border-radius: 16px !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05) !important;
        margin-bottom: 20px !important;
    }

    /* 3. 输入框圆角化与边框优化 */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 12px !important;
        border: 1px solid #E2E8F0 !important;
        background-color: #FFFFFF !important;
        padding: 10px !important;
    }

    /* 4. 侧边栏整体美化 */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #EDF2F7;
    }

    /* 5. 绿色按钮（确认/导出）深度定制：圆角与渐变效果 */
    div.stButton > button:first-child[kind="primary"] {
        width: 100% !important;
        background: linear-gradient(135deg, #28a745 0%, #218838 100%) !important;
        border: none !important;
        border-radius: 30px !important;
        padding: 14px 0 !important;
        font-weight: 700 !important;
        font-size: 1.1rem !important;
        box-shadow: 0 8px 15px rgba(40, 167, 69, 0.2) !important;
    }
    
    /* 6. 蓝色提示框圆角美化 */
    .stAlert {
        border-radius: 14px !important;
        border: none !important;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }

    /* 7. 手机端特定优化：压缩间距，调大字号 */
    @media (max-width: 768px) {
        .block-container { padding: 0.8rem 1rem !important; }
        h1 { font-size: 1.7rem !important; font-weight: 800 !important; margin-bottom: 0.5rem !important; }
        .stMarkdown p { font-size: 1rem !important; }
    }
</style>
""", unsafe_allow_html=True)

# --- 🔒 核心逻辑配置 ---
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

# 3. 侧边栏导航
with st.sidebar:
    st.header("⚙️ 体卫艺办公助手")
    st.success("● AI 核心已连接")
    st.caption("维护者：孙沛 | 龙华区教育局")
    st.divider()
    mode = st.radio("功能切换：", ["📝 领导公务单生成", "🔍 学校查号台"])
    
    if st.button("🔒 退出并锁定"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ----------------- 模块一：领导公务单生成器 -----------------
if mode == "📝 领导公务单生成":
    st.markdown("# 🚀 公务单自动生成")
    
    if st.session_state.step == 1:
        st.info("**💡 参考：** 明天上午10点在二楼多功能厅有个生涯教育座谈会，大概20人，孙沛对接，时长1小时，邀请灵芝主任参加")
        user_input = st.text_area("✍️ 请输入活动描述：", height=150, placeholder="在此输入文字或语音粘贴...", key="input_doc")
        
        if st.button("✨ 立即智能解析并生成", type="primary"):
            if not user_input:
                st.warning("请先输入内容")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                with st.spinner("AI 正在解析要素..."):
                    try:
                        full_prompt = f"你现在是龙华教育局资深笔杆子。请严格按 JSON 格式返回以下字段：title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。输入：{user_input}。今天是{datetime.now().strftime('%Y年%m月%d日')}。"
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct",
                            messages=[{"role": "user", "content": full_prompt}],
                            response_format={'type': 'json_object'}
                        )
                        st.session_state.parseddata_doc = json.loads(chat_completion.choices[0].message.content)
                        st.session_state.step = 2
                        st.rerun()
                    except Exception as e:
                        st.error(f"解析出错：{e}")

    elif st.session_state.step == 2:
        d = st.session_state.parseddata_doc
        with st.container():
            st.markdown("### 🧐 核心要素预览")
            t = st.text_input("📝 活动名称", d.get("title", ""))
            c = st.text_area("📄 申请理由/背景", d.get("content", ""), height=80)
            
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list): agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            if not agenda_val: agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("📋 详细议程", agenda_val, height=120)
            
            col1, col2 = st.columns(2)
            with col1:
                tm = st.text_input("⏰ 时间", d.get("time", ""))
                dr = st.text_input("⏳ 时长", d.get("duration", "1小时"))
            with col2:
                st.caption("时间可否调整：☑否")
                ct = st.text_input("👤 公务对接人", d.get("contact", "孙沛"))

            col3, col4, col5 = st.columns([2, 1, 1])
            with col3: pl = st.text_input("📍 地点", d.get("place", ""))
            with col4: nm = st.text_input("👥 人数", d.get("num", ""))
            with col5: pj = st.selectbox("📽️ 投影仪", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            dist_l = st.text_input("1. 拟请出席的区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("2. 拟请协调出席的局领导", d.get("bur_leader", ""))
            oth = st.text_input("建议参加单位", d.get("others") or "体卫艺劳科")

        col_back, col_down = st.columns([1, 2])
        with col_back:
            if st.button("⬅️ 返回重填"):
                st.session_state.step = 1
                st.rerun()

        with col_down:
            try:
                # 生成 Word 逻辑
                tpl = DocxTemplate("申报单模板.docx")
                tpl.render({"title":t,"content":c,"agenda":a,"time":tm,"duration":dr,"place":pl,"num":nm,"contact":ct,"projector":pj,"dist_leader":dist_l,"bur_leader":bur_l,"others":oth})
                bio = io.BytesIO()
                tpl.save(bio)

                # --- 沛沛专属命名逻辑 ---
                mmdd = datetime.now().strftime("%m%d") 
                raw_leader = dist_l.strip() if dist_l.strip() else bur_l.strip()
                if not raw_leader:
                    leader_name = "领导"
                else:
                    first = raw_leader.split('、')[0]
                    leader_name = "灵芝主任" if ("杨灵芝" in first or "灵芝" in first) else first
                
                # 最终文件名：1229-灵芝主任-体卫艺劳科-足球赛.docx
                final_filename = f"{mmdd}-{leader_name}-体卫艺劳科-{t}.docx"

                st.download_button(
                    label="💾 确认无误，导出 Word",
                    data=bio.getvalue(),
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
            except Exception as e:
                st.error(f"生成失败：{e}")

# ----------------- 模块二：龙华学校查号台 -----------------
elif mode == "🔍 学校查号台":
    st.markdown("# 🔍 学校查号台")
    if not st.session_state.contacts_authenticated:
        pwd = st.text_input("请输入授权密码：", type="password")
        if st.button("验证并进入", type="primary"):
            if pwd == CONTACT_PASSWORD:
                st.session_state.contacts_authenticated = True
                st.rerun()
            else: st.error("密码不正确")
    else:
        # 优化查号台 UI
        @st.cache_data
        def load_data():
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv').fillna('无')
        
        try:
            df = load_data()
            q = st.text_input("🔍 输入学校或人名：", placeholder="输入关键字...")
            
            if q:
                mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
                st.dataframe(df[mask], use_container_width=True, hide_index=True)
            else:
                st.write("📋 通讯录预览 (请输入关键词搜索)：")
                st.dataframe(df.head(5), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"通讯录加载失败，请检查 CSV 文件：{e}")
