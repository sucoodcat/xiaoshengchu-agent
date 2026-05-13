"""
小升初路径导航Agent v3.0 — Streamlit 网页应用
为昆明家长提供报名前/摇号前/录取后三阶段全流程导航
"""
import streamlit as st
from openai import OpenAI
from dotenv import load_dotenv
from pathlib import Path
from datetime import datetime, date
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import io
import os

# ── 环境加载 ──────────────────────────────────────────────
load_dotenv(Path(__file__).parent / ".env")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "deepseek-chat"

# ── 昆明学校清单（供下拉选择）─────────────────────────────
KUNMING_PRIMARY_SCHOOLS = {
    "五华区": ["武成小学", "春城小学", "红旗小学", "韶山小学", "龙翔小学", "文林小学",
               "莲华小学", "师大附小（文林校区）", "高新一小", "红云小学", "江岸小学"],
    "盘龙区": ["明通小学", "盘龙小学", "东华小学", "拓东小学", "金康园小学", "金星小学", "北京路小学"],
    "官渡区": ["南站小学", "关上实验学校", "东站小学", "民航路小学", "和平小学", "官渡中心学校"],
    "西山区": ["棕树营小学", "书林一小", "书林二小", "马街中心学校", "永昌小学"],
    "呈贡区": ["师大附小（呈贡校区）", "呈贡新区一小", "中华小学"],
    "其他": ["其他小学（不在列表中）"],
}

KUNMING_JUNIOR_SCHOOLS = {
    "民办-顶热（中签率10%-15%）": ["云大附中（一二一校区）", "云南师范大学附属实验中学", "长城中学"],
    "民办-热门（中签率15%-30%）": ["白塔中学", "滇池中学（昆三中分校）"],
    "民办-性价比（中签率25%-40%）": ["云子中学（昆十二中分校）", "师大附属润城学校"],
    "民转公/公参民": ["云大附中星耀校区"],
    "公办-五华区": ["昆八中", "昆二中", "昆二十四中", "昆十四中"],
    "公办-盘龙区": ["昆十中（白塔校区）", "昆十中（求实校区）", "昆明市实验中学"],
    "公办-官渡区": ["官渡区第五中学（昆铁三中）", "官渡区第一中学", "官渡区第二中学"],
    "公办-西山区": ["西山区第一中学", "昆五中"],
    "公办-呈贡区": ["师大附中呈贡校区", "呈贡一中", "昆三中呈贡校区"],
}

ALL_PRIMARY = [f"{s}（{d}）" for d, sl in KUNMING_PRIMARY_SCHOOLS.items() for s in sl]
ALL_JUNIOR = [f"{s} [{c}]" for c, sl in KUNMING_JUNIOR_SCHOOLS.items() for s in sl]

# ── 阶段选项与学生标签 ──────────────────────────────────
PHASE_OPTIONS = [
    "自动判定（根据当前日期）",
    "报名前 — 我还没选好学校，需要择校分析",
    "报名后-摇号前 — 已报名，等待摇号，需要行动规划",
    "录取后 — 已确定学校，需要初中入学准备",
]

STUDENT_PERSONALITY = [
    "开朗外向", "安静内向", "沉稳细致", "活泼好动",
    "自律自觉", "需要督促", "好奇心强", "有主见",
]

STUDENT_INTERESTS = [
    "热爱运动", "喜欢读书", "擅长主持/演讲", "喜欢科学实验",
    "喜欢编程/计算机", "喜欢音乐", "喜欢绘画/美术",
    "喜欢写作", "擅长手工/制作", "喜欢小动物/自然",
]


# ── 页面配置 ──────────────────────────────────────────────
st.set_page_config(
    page_title="小升初路径导航 v3 · 昆明",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── 自定义CSS ─────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #f5f7fa 0%, #e8ecf1 100%); }
    .main-header {
        background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
        color: white; padding: 20px 28px; border-radius: 14px;
        margin-bottom: 20px; box-shadow: 0 4px 16px rgba(26,115,232,0.25);
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
    .main-header p { margin: 4px 0 0; opacity: 0.9; font-size: 0.85rem; }
    .form-card {
        background: white; border-radius: 14px; padding: 24px 28px;
        margin: 12px 0; box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    }
    .form-section-title {
        font-size: 1rem; font-weight: 700; color: #1a73e8;
        margin-bottom: 12px; padding-bottom: 6px; border-bottom: 2px solid #e8ecf1;
    }
    .scenario-card-a {
        background: #e8f5e9; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 4px solid #4caf50;
    }
    .scenario-card-b {
        background: #fff3e0; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 4px solid #ff9800;
    }
    .countdown {
        background: #1a73e8; color: white; padding: 4px 12px;
        border-radius: 20px; font-weight: 600; font-size: 0.85rem;
        display: inline-block;
    }
    .info-summary {
        background: white; border-radius: 10px; padding: 12px 16px;
        margin: 8px 0; box-shadow: 0 1px 6px rgba(0,0,0,0.05);
        font-size: 0.85rem;
    }
    .info-summary b { color: #1a73e8; }
    .stButton > button {
        background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
        color: white; border: none; border-radius: 10px;
        padding: 12px 32px; font-size: 1.05rem; font-weight: 700;
        width: 100%; transition: all 0.2s;
    }
    .stButton > button:hover {
        box-shadow: 0 4px 14px rgba(26,115,232,0.4);
        transform: translateY(-1px);
    }
    /* 侧边栏精简 */
    .stSidebar { background: white; padding-top: 16px; }
    @media (max-width: 768px) {
        .main-header h1 { font-size: 1.2rem; }
    }
</style>
""", unsafe_allow_html=True)


# ── 知识库加载（缓存）────────────────────────────────────
@st.cache_resource
def load_system_prompt():
    root = Path(__file__).parent
    parts = []
    claude_md = root / "CLAUDE.md"
    if claude_md.exists():
        parts.append(claude_md.read_text(encoding="utf-8"))
    kb_dir = root / "knowledge-base"
    if kb_dir.exists():
        parts.append("\n\n# 以下为知识库数据，请严格引用：\n")
        for kb_file in sorted(kb_dir.glob("*.md")):
            parts.append(f"\n## {kb_file.stem}\n")
            parts.append(kb_file.read_text(encoding="utf-8"))
    today = date.today()
    parts.append(f"\n\n# 当前日期：{today}（{today.strftime('%Y年%m月%d日')}）")
    parts.append("\n请在回答中基于当前日期计算倒计时，区分已完成/进行中/即将到来的事项。")
    return "\n".join(parts)


def get_timeline_context():
    today = date.today()
    milestones = [
        ("志愿填报截止", date(2026, 5, 31)),
        ("电脑摇号公布", date(2026, 6, 23)),
        ("补录窗口期", date(2026, 6, 24)),
        ("公办录取公布", date(2026, 7, 3)),
        ("分班考试季开始", date(2026, 7, 25)),
        ("新生报到", date(2026, 8, 25)),
        ("正式开学", date(2026, 9, 1)),
    ]
    lines = ["## 当前时间节点状态（基于今日日期自动计算）", ""]
    for name, dt in milestones:
        delta = (dt - today).days
        if delta < 0:
            lines.append(f"- ✅ ~~{name}~~（已过去 {-delta} 天）")
        elif delta == 0:
            lines.append(f"- 🔴 **{name}** — 就是今天！")
        elif delta <= 7:
            lines.append(f"- 🔴 **{name}** — 倒计时 **{delta} 天**")
        elif delta <= 30:
            lines.append(f"- 🟡 {name} — 还有 {delta} 天")
        else:
            lines.append(f"- ⏳ {name} — 还有 {delta} 天")
    return "\n".join(lines)


def build_user_context():
    ui = st.session_state.user_info
    ctx = "## 用户填写的信息\n"
    ctx += f"- 升学阶段：{ui.get('phase') or '未指定'}\n"
    ctx += f"- 孩子姓名：{ui.get('child_name') or '未提供'}\n"
    ctx += f"- 性别：{ui.get('gender') or '未提供'}\n"
    ctx += f"- 就读小学：{ui.get('primary_school') or '未提供'}\n"
    ctx += f"- 第一志愿：{ui.get('target_school_1') or '未提供'}\n"
    if ui.get('target_school_2'):
        ctx += f"- 第二志愿/备选：{ui['target_school_2']}\n"
    if ui.get('target_school_3'):
        ctx += f"- 第三志愿/备选：{ui['target_school_3']}\n"
    ctx += f"- 学业水平：{ui.get('academic_level') or '未提供'}\n"
    ctx += f"- 三好学生/荣誉级别：{ui.get('honor_level') or '无'}"
    if ui.get('honor_detail'):
        ctx += f"（{ui['honor_detail']}）"
    ctx += "\n"
    ctx += f"- 学科竞赛获奖：{ui.get('competition_detail') or '无'}\n"
    ctx += f"- 体育特长级别：{ui.get('sports_level') or '无'}"
    if ui.get('sports_detail'):
        ctx += f" — 项目：{ui['sports_detail']}"
    ctx += "\n"
    ctx += f"- 艺术特长级别：{ui.get('art_level') or '无'}"
    if ui.get('art_detail'):
        ctx += f" — 项目：{ui['art_detail']}"
    ctx += "\n"
    p_tags = list(ui.get('personality_tags', []))
    if ui.get('personality_custom'):
        p_tags.extend([t.strip() for t in ui['personality_custom'].split(',') if t.strip()])
    ctx += f"- 性格标签：{', '.join(p_tags) if p_tags else '未提供'}\n"
    i_tags = list(ui.get('interest_tags', []))
    if ui.get('interest_custom'):
        i_tags.extend([t.strip() for t in ui['interest_custom'].split(',') if t.strip()])
    ctx += f"- 兴趣标签：{', '.join(i_tags) if i_tags else '未提供'}\n"
    ctx += f"- 户籍：{ui.get('household') or '未提供'}\n"
    ctx += f"- 双胞胎/多胞胎：{'是' if ui.get('is_twins') else '否'}\n"
    ctx += f"- 家迁/搬迁计划：{ui.get('relocation') or '无'}\n"
    ctx += f"- 家迁时间：{ui.get('relocation_time') or '未提供'}\n"
    ctx += f"- 手机号：{ui.get('phone') or '未提供'}\n"
    ctx += f"- 其他补充：{ui.get('extra_info') or '未提供'}\n"
    if ui.get('resume_text'):
        ctx += f"\n## 学生简历内容（家长上传）\n{ui['resume_text']}\n"
    return ctx


# ── 初始化 Session State ─────────────────────────────────
default_user_info = {
    "child_name": "", "gender": "",
    "phase": PHASE_OPTIONS[0],
    "primary_school": "", "target_school_1": "",
    "target_school_2": "", "target_school_3": "",
    "academic_level": "",
    "honor_level": "", "honor_detail": "",
    "competition_detail": "",
    "sports_level": "", "sports_detail": "",
    "art_level": "", "art_detail": "",
    "personality_tags": [], "personality_custom": "",
    "interest_tags": [], "interest_custom": "",
    "household": "昆明主城区", "is_twins": False,
    "relocation": "", "relocation_time": "",
    "phone": "", "extra_info": "",
    "resume_text": "",
}

if "messages" not in st.session_state:
    st.session_state.messages = []
if "user_info" not in st.session_state:
    st.session_state.user_info = dict(default_user_info)
if "form_submitted" not in st.session_state:
    st.session_state.form_submitted = False


def reset_all():
    st.session_state.messages = []
    st.session_state.user_info = dict(default_user_info)
    st.session_state.form_submitted = False
    st.rerun()


def submit_form():
    ui = st.session_state.user_info
    if not ui["primary_school"]:
        st.error("请至少填写「就读小学」。")
        st.stop()

    # 根据阶段生成不同的初始请求
    phase = ui.get("phase", "")
    if "报名前" in phase:
        prompt = (
            "请根据我填写的信息，生成「阶段一：择校分析报告」。"
            "包括：学生画像速览、对口公办分析、民办学校推荐与对比、择校决策建议、报名须知。"
            "如果第一志愿已填，把它作为重点分析对象。不要输出双场景日历。"
        )
    elif "录取后" in phase:
        prompt = (
            "请根据我填写的信息，生成「阶段三：初中入学准备指南」。"
            "包括：录取结果确认、新生报到清单、分班考试备考方案、小初衔接学习建议、初中三年关键节点、初升高远景。"
            "不要输出双场景日历。"
        )
    else:
        prompt = (
            "请根据我填写的信息，生成「阶段二：摇号前行动规划」。"
            "包括：选择确认、当前状态与倒计时、当前应做事项（紧急度排列）、"
            "☀️摇中行动日历、🌧️未摇中应对方案、初升高远景。"
            "必须同时展示摇中和未摇中两种场景。"
        )

    full = f"{prompt}\n\n{build_user_context()}"
    st.session_state.messages.append({"role": "user", "content": full})
    st.session_state.form_submitted = True
    st.rerun()


# ── 侧边栏（精简）─────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🎓 小升初路径导航")
    st.caption("昆明市 · v3.0")
    st.divider()
    st.button("🔄 重新开始", on_click=reset_all, use_container_width=True)
    st.divider()
    with st.expander("🔒 隐私与声明"):
        st.caption(
            "• 数据仅用于生成报告\n"
            "• 脱敏后用于宏观分析\n"
            "• 绝不泄露给第三方\n"
            "• API Key由系统统一管理\n"
            "• 无需家长输入任何密钥"
        )
    st.caption("⚠️ 信息整合和路径推演仅供参考，不构成决策依据。")


# ── 主界面 ────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🎓 小升初路径导航助手</h1>
    <p>昆明市 · 报名前/摇号前/录取后 三阶段全流程导航 · v3.0</p>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# 阶段1：未提交表单 → 显示信息填写界面
# ═══════════════════════════════════════════════════════════
if not st.session_state.form_submitted:
    ui = st.session_state.user_info

    st.markdown("""
    <div style="max-width:900px; margin:0 auto;">
    <p style="font-size:1.05rem; color:#555; margin-bottom:16px;">
    请填写以下信息后点击<b>「开始规划」</b>。<span style="color:#d32f2f;">*</span> 为必填项，其他选填项越详细，规划越精准。
    </p>
    </div>
    """, unsafe_allow_html=True)

    # ── 阶段选择 ──
    with st.container():
        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        st.markdown('<p class="form-section-title">📌 升学阶段</p>', unsafe_allow_html=True)
        ui["phase"] = st.selectbox(
            "您当前处于哪个阶段？",
            PHASE_OPTIONS,
            index=PHASE_OPTIONS.index(ui["phase"]) if ui["phase"] in PHASE_OPTIONS else 0,
            key="f_phase",
            help="选择后将生成对应阶段的专属报告。默认自动根据日期判定。"
        )
        st.caption("报名前→择校分析 | 报名后-摇号前→行动规划 | 录取后→入学准备")
        st.markdown("</div>", unsafe_allow_html=True)

    # ── 表单卡片 ──
    with st.container():
        st.markdown('<div class="form-card">', unsafe_allow_html=True)

        # --- 第1行：基本信息 ---
        st.markdown('<p class="form-section-title">👤 基本信息</p>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([2, 1, 2])
        with c1:
            ui["child_name"] = st.text_input(
                "孩子姓名（选填）", value=ui["child_name"],
                placeholder="小名或昵称也可以", key="f_name"
            )
        with c2:
            ui["gender"] = st.radio(
                "性别", ["", "男", "女"],
                index=0 if not ui["gender"] else (1 if ui["gender"] == "男" else 2),
                horizontal=True, key="f_gender"
            )
        with c3:
            ui["phone"] = st.text_input(
                "手机号（选填，用于关键提醒）", value=ui["phone"],
                placeholder="不验证，纯文本收集", key="f_phone"
            )

        st.markdown("</div>", unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="form-card">', unsafe_allow_html=True)

        # --- 第2行：学校信息 ---
        st.markdown('<p class="form-section-title">🏫 学校信息</p>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            ui["primary_school"] = st.selectbox(
                "就读小学 *", [""] + ALL_PRIMARY,
                index=0 if not ui["primary_school"]
                      else (ALL_PRIMARY.index(ui["primary_school"]) + 1
                            if ui["primary_school"] in ALL_PRIMARY else 0),
                key="f_ps"
            )
        with c2:
            ui["household"] = st.selectbox(
                "户籍类型", ["昆明主城区", "昆明郊县", "云南省内其他", "外省"],
                index=["昆明主城区", "昆明郊县", "云南省内其他", "外省"].index(ui["household"])
                if ui["household"] in ["昆明主城区", "昆明郊县", "云南省内其他", "外省"] else 0,
                key="f_hh"
            )

        st.markdown("**拟报初中**（第一志愿必填，2-3为备选）")
        c1, c2, c3 = st.columns(3)
        with c1:
            ui["target_school_1"] = st.selectbox(
                "第一志愿 *", [""] + ALL_JUNIOR,
                index=0 if not ui["target_school_1"]
                      else (ALL_JUNIOR.index(ui["target_school_1"]) + 1
                            if ui["target_school_1"] in ALL_JUNIOR else 0),
                key="f_ts1"
            )
        with c2:
            ui["target_school_2"] = st.selectbox(
                "第二志愿/备选1", [""] + ALL_JUNIOR,
                index=0 if not ui["target_school_2"]
                      else (ALL_JUNIOR.index(ui["target_school_2"]) + 1
                            if ui["target_school_2"] in ALL_JUNIOR else 0),
                key="f_ts2"
            )
        with c3:
            ui["target_school_3"] = st.selectbox(
                "第三志愿/备选2", [""] + ALL_JUNIOR,
                index=0 if not ui["target_school_3"]
                      else (ALL_JUNIOR.index(ui["target_school_3"]) + 1
                            if ui["target_school_3"] in ALL_JUNIOR else 0),
                key="f_ts3"
            )

        st.markdown("</div>", unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="form-card">', unsafe_allow_html=True)

        # --- 第3行：荣誉与特长 ---
        st.markdown('<p class="form-section-title">🏆 荣誉与特长（有则填写，越详细规划越精准）</p>', unsafe_allow_html=True)

        # 三好学生
        c1, c2 = st.columns([1, 2])
        with c1:
            ui["honor_level"] = st.selectbox(
                "三好学生/优秀学生干部",
                ["无", "校级", "区级", "市级", "省级"],
                index=["无", "校级", "区级", "市级", "省级"].index(ui["honor_level"])
                if ui["honor_level"] in ["无", "校级", "区级", "市级", "省级"] else 0,
                key="f_hl"
            )
        with c2:
            ui["honor_detail"] = st.text_input(
                "荣誉具体描述（选填）", value=ui["honor_detail"],
                placeholder="如：2025年昆明市三好学生、校级优秀学生干部", key="f_hd"
            )

        # 学科竞赛
        ui["competition_detail"] = st.text_input(
            "学科竞赛获奖", value=ui["competition_detail"],
            placeholder="如：华杯省一等奖、NOIP普及组二等奖、外研社杯省一等奖", key="f_comp"
        )

        # 体育特长
        c1, c2 = st.columns([1, 2])
        with c1:
            ui["sports_level"] = st.selectbox(
                "体育特长级别",
                ["无", "校级", "区/县级", "市级", "省级", "国家级"],
                index=["无", "校级", "区/县级", "市级", "省级", "国家级"].index(ui["sports_level"])
                if ui["sports_level"] in ["无", "校级", "区/县级", "市级", "省级", "国家级"] else 0,
                key="f_sl"
            )
        with c2:
            ui["sports_detail"] = st.text_input(
                "体育特长项目", value=ui["sports_detail"],
                placeholder="如：篮球（校队主力）、田径200m（市级第二）", key="f_sd"
            )

        # 艺术特长
        c1, c2 = st.columns([1, 2])
        with c1:
            ui["art_level"] = st.selectbox(
                "艺术特长级别",
                ["无", "校级", "区/县级", "市级", "省级", "国家级"],
                index=["无", "校级", "区/县级", "市级", "省级", "国家级"].index(ui["art_level"])
                if ui["art_level"] in ["无", "校级", "区/县级", "市级", "省级", "国家级"] else 0,
                key="f_alvl"
            )
        with c2:
            ui["art_detail"] = st.text_input(
                "艺术特长项目", value=ui["art_detail"],
                placeholder="如：钢琴十级、管乐团长笛首席", key="f_ad"
            )

        st.markdown("</div>", unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="form-card">', unsafe_allow_html=True)

        # --- 第4行：学业与家庭情况 ---
        st.markdown('<p class="form-section-title">🌱 学业与家庭情况</p>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            ui["academic_level"] = st.selectbox(
                "学业水平自评",
                ["", "优秀（班级前10%）", "良好（班级前30%）", "中等", "待提升"],
                index=0 if not ui["academic_level"]
                      else ["", "优秀（班级前10%）", "良好（班级前30%）", "中等", "待提升"].index(ui["academic_level"])
                      if ui["academic_level"] in ["", "优秀（班级前10%）", "良好（班级前30%）", "中等", "待提升"] else 0,
                key="f_al"
            )
        with c2:
            ui["is_twins"] = st.checkbox("双胞胎/多胞胎", value=ui["is_twins"], key="f_twins")
        with c3:
            ui["household"] = st.selectbox(
                "户籍类型", ["昆明主城区", "昆明郊县", "云南省内其他", "外省"],
                index=["昆明主城区", "昆明郊县", "云南省内其他", "外省"].index(ui["household"])
                if ui["household"] in ["昆明主城区", "昆明郊县", "云南省内其他", "外省"] else 0,
                key="f_hh2"
            )

        # 家迁
        c1, c2 = st.columns([2, 1])
        with c1:
            ui["relocation"] = st.text_area(
                "家迁/搬迁计划", value=ui["relocation"],
                placeholder="如：已搬到呈贡区XX小区 / 计划7月前迁到盘龙区 / 工作调动可能离开昆明",
                height=60, key="f_reloc"
            )
        with c2:
            ui["relocation_time"] = st.selectbox(
                "搬迁时间节点",
                ["暂无搬迁", "已搬迁", "1个月内", "3个月内", "半年内", "计划中/时间未定"],
                index=["暂无搬迁", "已搬迁", "1个月内", "3个月内", "半年内", "计划中/时间未定"].index(ui["relocation_time"])
                if ui["relocation_time"] in ["暂无搬迁", "已搬迁", "1个月内", "3个月内", "半年内", "计划中/时间未定"] else 0,
                key="f_rt"
            )

        ui["extra_info"] = st.text_area(
            "其他补充信息", value=ui["extra_info"],
            placeholder="任何您觉得重要的：对学校的特殊要求、家庭特殊情况等",
            height=60, key="f_extra"
        )

        # 性格与兴趣标签
        st.markdown("**孩子性格与兴趣**（可多选，也可自行填写）")
        c1, c2 = st.columns(2)
        with c1:
            ui["personality_tags"] = st.multiselect(
                "性格特点", STUDENT_PERSONALITY,
                default=ui.get("personality_tags", []) if isinstance(ui.get("personality_tags"), list) else [],
                key="f_ptags"
            )
            ui["personality_custom"] = st.text_input(
                "自定义性格标签（逗号分隔）",
                value=ui.get("personality_custom", ""),
                placeholder="如：乐于助人、有领导力、专注力强",
                key="f_pcustom"
            )
        with c2:
            ui["interest_tags"] = st.multiselect(
                "兴趣爱好", STUDENT_INTERESTS,
                default=ui.get("interest_tags", []) if isinstance(ui.get("interest_tags"), list) else [],
                key="f_itags"
            )
            ui["interest_custom"] = st.text_input(
                "自定义兴趣爱好（逗号分隔）",
                value=ui.get("interest_custom", ""),
                placeholder="如：国际象棋、机器人搭建、观鸟",
                key="f_icustom"
            )

        st.markdown("</div>", unsafe_allow_html=True)

    # --- 简历上传 ---
    with st.container():
        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        st.markdown('<p class="form-section-title">📄 学生简历上传（选填）</p>', unsafe_allow_html=True)
        st.caption("支持 PDF、Word(.docx)、图片(.png/.jpg)。简历中的信息将自动提取并用于分析。")

        uploaded_file = st.file_uploader(
            "选择文件", type=["pdf", "docx", "png", "jpg", "jpeg"], key="f_resume",
        )
        if uploaded_file is not None:
            try:
                file_type = uploaded_file.name.split(".")[-1].lower()
                resume_text = ""

                if file_type == "pdf":
                    pdf_reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
                    for page in pdf_reader.pages:
                        text = page.extract_text()
                        if text:
                            resume_text += text + "\n"

                elif file_type == "docx":
                    doc = Document(io.BytesIO(uploaded_file.getvalue()))
                    for para in doc.paragraphs:
                        if para.text.strip():
                            resume_text += para.text + "\n"
                    # Also extract tables
                    for table in doc.tables:
                        for row in table.rows:
                            row_text = " | ".join(cell.text for cell in row.cells)
                            if row_text.strip():
                                resume_text += row_text + "\n"

                elif file_type in ("png", "jpg", "jpeg"):
                    resume_text = (
                        "[图片简历已上传。由于图片格式无法自动提取文字，"
                        "请在下方「其他补充信息」中手动输入简历关键内容，"
                        "或上传PDF/Word版本的简历。]"
                    )

                ui["resume_text"] = resume_text.strip()
                if ui["resume_text"]:
                    extracted_len = len(ui["resume_text"])
                    if file_type in ("png", "jpg", "jpeg"):
                        st.warning("⚠️ 图片格式无法自动提取文字，建议上传PDF或Word版本。")
                    else:
                        st.success(f"✅ 简历解析成功，共提取约 {extracted_len} 字")
                    with st.expander("查看解析内容"):
                        st.text(ui["resume_text"][:2000])
            except Exception as e:
                st.warning(f"简历解析异常：{e}。您可将简历内容粘贴到上方「其他补充信息」中。")
                ui["resume_text"] = ""

        st.markdown("</div>", unsafe_allow_html=True)

    # --- 提交按钮 ---
    st.markdown('<div style="max-width:900px; margin:0 auto;">', unsafe_allow_html=True)
    st.button("🚀 开始规划", on_click=submit_form, use_container_width=True)

    # 时间线速览
    today = date.today()
    days_left = (date(2026, 6, 23) - today).days
    st.caption(
        f"📅 今日：{today.strftime('%Y年%m月%d日')} ｜ "
        f"⏳ 距离摇号公布还有 {days_left} 天 ｜ "
        f"🔴 志愿填报截止还有 {(date(2026, 5, 31) - today).days} 天"
    )
    st.markdown("</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# 阶段2：已提交 → 显示结果 + 信息摘要 + 追问入口
# ═══════════════════════════════════════════════════════════
else:
    # ── 信息摘要卡 ──
    ui = st.session_state.user_info
    with st.expander("📋 已填写的信息（点击展开查看/修改）", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.caption(f"**阶段**：{ui.get('phase','—')}")
            st.caption(f"**姓名**：{ui['child_name'] or '—'}")
            st.caption(f"**性别**：{ui['gender'] or '—'}")
            st.caption(f"**双胞胎**：{'是' if ui['is_twins'] else '否'}")
        with c2:
            st.caption(f"**小学**：{ui['primary_school'] or '—'}")
            st.caption(f"**户籍**：{ui['household']}")
            st.caption(f"**学业水平**：{ui['academic_level'] or '—'}")
        with c3:
            st.caption(f"**第一志愿**：{ui['target_school_1'] or '—'}")
            st.caption(f"**备选1**：{ui['target_school_2'] or '—'}")
            st.caption(f"**备选2**：{ui['target_school_3'] or '—'}")
        with c4:
            st.caption(f"**三好学生**：{ui.get('honor_level','无')}")
            st.caption(f"**竞赛**：{ui.get('competition_detail','—') or '—'}")
            st.caption(f"**体育**：{ui.get('sports_level','无')} {ui.get('sports_detail','')}")
            st.caption(f"**艺术**：{ui.get('art_level','无')} {ui.get('art_detail','')}")
        ptags = ui.get('personality_tags', [])
        itags = ui.get('interest_tags', [])
        if ptags or itags:
            st.caption(f"**性格**：{', '.join(ptags) if ptags else '—'} ｜ **兴趣**：{', '.join(itags) if itags else '—'}")
        if ui.get("relocation"):
            st.caption(f"**家迁**：{ui['relocation']}（{ui.get('relocation_time','')}）")
        if ui.get("honor_detail"):
            st.caption(f"**荣誉详情**：{ui['honor_detail']}")
        if ui.get("extra_info"):
            st.caption(f"**补充**：{ui['extra_info']}")
        if ui.get("resume_text"):
            st.caption(f"**简历**：已上传（{len(ui['resume_text'])}字）")

        if st.button("✏️ 修改信息，重新规划", use_container_width=False):
            st.session_state.form_submitted = False
            st.session_state.messages = []
            st.rerun()

    st.divider()

    # ── 结果展示区 ──
    # 如果尚未调用API（刚提交form），自动触发API调用
    if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
        last_user_msg = st.session_state.messages[-1]["content"]

        # 检查是否已经在处理中（避免重复调用）
        # 如果最后一条不是assistant，则需要调用
        need_api_call = True
        if len(st.session_state.messages) >= 2 and st.session_state.messages[-1]["role"] == "assistant":
            need_api_call = False

        if need_api_call:
            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = ""

                if not DEEPSEEK_API_KEY or DEEPSEEK_API_KEY == "sk-your-deepseek-key-here":
                    full_response = (
                        "❌ **服务器未配置有效的 DeepSeek API Key。**\n\n"
                        "请联系管理员编辑 `.env` 文件，将 `DEEPSEEK_API_KEY` 设置为真实Key。\n"
                        "获取地址：[platform.deepseek.com](https://platform.deepseek.com)\n\n"
                        "配置完成后重启应用即可。"
                    )
                    message_placeholder.markdown(full_response)
                else:
                    try:
                        client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)
                        system_prompt = load_system_prompt()
                        timeline_ctx = get_timeline_context()
                        full_system = system_prompt + "\n\n" + timeline_ctx

                        api_messages = [{"role": "system", "content": full_system}]
                        for m in st.session_state.messages:
                            role = m["role"]
                            content = m["content"]
                            api_messages.append({"role": role, "content": content})

                        stream = client.chat.completions.create(
                            model=DEFAULT_MODEL,
                            messages=api_messages,
                            max_tokens=4096,
                            temperature=0.7,
                            stream=True,
                        )
                        for chunk in stream:
                            if chunk.choices[0].delta.content:
                                full_response += chunk.choices[0].delta.content
                                message_placeholder.markdown(full_response + " ▌")
                        message_placeholder.markdown(full_response)

                    except Exception as e:
                        err = str(e)
                        if "authentication" in err.lower() or "api_key" in err.lower():
                            full_response = "❌ **API Key 无效**。请联系管理员检查 `.env` 文件。"
                        elif "rate" in err.lower():
                            full_response = "⏳ **请求过于频繁**，请稍等几秒后再试。"
                        else:
                            full_response = f"❌ 出错了：{err}"
                        message_placeholder.markdown(full_response)

                st.session_state.messages.append({"role": "assistant", "content": full_response})
                st.rerun()

    # ── 展示已有的对话历史 ──
    for msg in st.session_state.messages:
        role_label = "user" if msg["role"] == "user" else "assistant"
        with st.chat_message(role_label):
            # 用户消息中剥离注入的上下文，只显示用户可见内容
            if msg["role"] == "user":
                display = msg["content"].split("\n\n## 用户填写的信息")[0]
                st.markdown(display)
            else:
                st.markdown(msg["content"])

    st.divider()

    # ── 追问入口 ──
    st.markdown("### 💬 继续咨询")
    st.caption("输入「**对比**」比较志愿学校 · 输入「**分班考**」了解备考详情 · 或直接输入你的问题")

    if followup := st.chat_input("输入追问内容..."):
        if not DEEPSEEK_API_KEY or DEEPSEEK_API_KEY == "sk-your-deepseek-key-here":
            st.error("❌ 服务器未配置有效的 DeepSeek API Key。请联系管理员。")
            st.stop()

        st.session_state.messages.append({
            "role": "user",
            "content": f"{followup}\n\n{build_user_context()}"
        })
        st.rerun()
