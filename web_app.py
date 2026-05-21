"""
培正学院毕业季超级工具箱 - 网页 Web 版 (Streamlit 驱动)
用法：streamlit run web_app.py
"""
import streamlit as st
import os
import tempfile
from Format_verify_tool import analyze_proposal, HTMLReporter
from Thesis_verify_tool import ThesisFormatVerifier

# ─────────────────────────────────────────────────────────────
# 页面基础配置
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="培正学院 - 毕业季超级工具箱",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# 自定义 CSS 注入
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* 隐藏默认 header/footer 水印 */
    header[data-testid="stHeader"] { display: none; }
    footer { visibility: hidden; }

    /* 全局字体优化 */
    .stApp {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans SC", sans-serif;
    }

    /* 侧边栏美化 */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    }

    /* 侧边栏全部文字强制提亮 */
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] .stMarkdown {
        color: #F8FAFC !important;
        text-shadow: 0px 1px 2px rgba(0, 0, 0, 0.3) !important;
    }

    /* Radio 选项：选中态高亮 */
    section[data-testid="stSidebar"] .stRadio label span {
        color: #e2e8f0 !important;
        font-size: 15px;
        transition: all 0.2s ease-in-out;
    }
    section[data-testid="stSidebar"] .stRadio [data-baseweb="radio"] div[role="radio"][aria-checked="true"] {
        background-color: #6366f1 !important;
        border-color: #6366f1 !important;
    }

    /* Radio 悬浮：呼吸感微微放大 */
    [data-testid="stSidebar"] .stRadio label {
        transition: all 0.2s ease-in-out;
        transform-origin: left center;
        padding: 2px 0;
        border-radius: 6px;
    }
    [data-testid="stSidebar"] .stRadio label:hover {
        transform: scale(1.03);
    }
    [data-testid="stSidebar"] .stRadio label:hover span {
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }

    /* GitHub 链接：亮蓝色 + 悬浮呼吸放大 */
    [data-testid="stSidebar"] a {
        color: #93C5FD !important;
        font-weight: bold !important;
        text-decoration: none !important;
        transition: all 0.2s ease-in-out;
        display: inline-block;
    }
    [data-testid="stSidebar"] a:hover {
        color: #BFDBFE !important;
        text-decoration: underline !important;
        transform: scale(1.03);
    }

    /* Metric 卡片美化 */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 12px;
        padding: 16px 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    [data-testid="stMetric"] label {
        color: #475569 !important;
        font-weight: 600 !important;
    }

    /* Expander 样式微调 */
    details[data-testid="stExpander"] {
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        overflow: hidden;
    }

    /* 按钮圆角 */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }

    /* 顶部标题区 */
    .hero-banner {
        text-align: center;
        padding: 20px 0 10px;
    }
    .hero-banner h1 {
        font-size: 32px;
        background: linear-gradient(135deg, #6366f1, #8b5cf6, #a855f7);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 4px;
    }
    .hero-banner p {
        color: #94a3b8;
        font-size: 14px;
        margin: 0;
    }

    /* 自定义 Warning Box */
    .custom-warning {
        background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
        border-left: 5px solid #f59e0b;
        border-radius: 8px;
        padding: 14px 18px;
        margin: 12px 0;
        font-size: 14px;
        color: #92400e;
        line-height: 1.6;
    }
    .custom-warning strong {
        color: #b45309;
    }

    /* 通过/失败结果卡片 */
    .result-pass {
        background: #f0fdf4;
        border-left: 4px solid #22c55e;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0;
    }
    .result-fail {
        background: #fef2f2;
        border-left: 4px solid #ef4444;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# 侧边栏导航
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        "<h2 style='text-align:center; color:#a5b4fc; margin-bottom:4px;'>🎓 毕业季超级工具箱</h2>"
        "<p style='text-align:center; color:#64748b; font-size:12px; margin-bottom:24px;'>"
        "培正学院 · 格式校验双引擎</p>",
        unsafe_allow_html=True,
    )
    page = st.radio(
        "选择功能",
        ["📝 开题报告格式校验", "📑 毕业论文终稿查验"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown(
        "<p style='text-align:center; color:#475569; font-size:12px; line-height:1.8;'>"
        "Developed with ❤️ by 林格<br>"
        "<a href='https://github.com/linjunhao024-byte/The_All_Web' "
        "target='_blank' style='color:#818cf8; text-decoration:none;'>"
        "GitHub 开源地址</a></p>",
        unsafe_allow_html=True,
    )


# ═════════════════════════════════════════════════════════════
# 模块 1：开题报告格式校验
# ═════════════════════════════════════════════════════════════
def render_proposal_checker():
    # Hero Banner
    st.markdown(
        "<div class='hero-banner'>"
        "<h1>📝 开题报告格式校验助手</h1>"
        "<p>一键解析 · 多维交叉校对 · 精美诊断报告</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.divider()

    # 参数配置
    st.markdown("#### ⚙️ 高级自定义参数配置（可选）")
    col1, col2 = st.columns(2)
    with col1:
        line_spacing = st.number_input(
            "预期行距", min_value=1.0, max_value=3.0, value=1.5, step=0.05
        )
        min_refs = st.number_input(
            "文献数量下限 (篇)", min_value=1, max_value=50, value=10, step=1
        )
    with col2:
        min_words = st.number_input(
            "综述字数下限", min_value=500, max_value=10000, value=1800, step=100
        )
        max_words = st.number_input(
            "综述字数上限", min_value=500, max_value=20000, value=2200, step=100
        )

    st.markdown("##### 🛠️ 规则开关控制")
    c1, c2, c3 = st.columns(3)
    with c1:
        check_tutor_space = st.checkbox("导师职称空格规范", value=True)
    with c2:
        check_indent = st.checkbox("参考文献悬挂缩进", value=True)
    with c3:
        check_timeline = st.checkbox("进度安排时间线顺叙", value=True)

    st.divider()

    # 文件上传
    st.markdown("#### 📁 上传你的开题报告")
    st.markdown(
        "<div class='custom-warning'>"
        "<strong>⚠️ 免责声明：</strong>"
        "本工具格式校验结果仅供参考。由于文档结构可能存在原生性错乱，"
        "实际通过标准请以导师最终意见为准，建议生成报告后人工复查一遍。"
        "</div>",
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "仅支持 .docx 格式的 Word 文档", type=["docx"], key="proposal_uploader"
    )

    if uploaded_file is not None:
        custom_config = {
            "line_spacing": line_spacing,
            "min_refs": min_refs,
            "min_words": min_words,
            "max_words": max_words,
            "check_tutor_space": check_tutor_space,
            "check_indent": check_indent,
            "check_timeline": check_timeline,
        }

        if st.button(
            "🚀 开始一键格式校验", type="primary", use_container_width=True, key="proposal_btn"
        ):
            with st.spinner("🔍 正在调用底层解析引擎，执行多维矩阵交叉校对..."):
                try:
                    with tempfile.NamedTemporaryFile(
                        delete=False, suffix=".docx"
                    ) as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name

                    result = analyze_proposal(tmp_path, config=custom_config)
                    doc_name = uploaded_file.name
                    reporter = HTMLReporter(result, doc_name=doc_name)
                    report_path = reporter.generate("web_report.html")

                    with open(report_path, "r", encoding="utf-8") as f:
                        html_content = f.read()

                    os.remove(tmp_path)
                    os.remove(report_path)

                    st.balloons()
                    st.success("🎉 校验完成！诊断报告已在下方生成。")
                    st.components.v1.html(html_content, height=900, scrolling=True)

                except Exception as e:
                    st.error(f"❌ 校验过程中发生错误，请检查文档是否损坏。错误信息：{e}")


# ═════════════════════════════════════════════════════════════
# 模块 2：毕业论文终稿查验
# ═════════════════════════════════════════════════════════════
def render_thesis_checker():
    # Hero Banner
    st.markdown(
        "<div class='hero-banner'>"
        "<h1>📑 毕业论文终稿查验</h1>"
        "<p>页边距 · 字体字号 · 标题层级 · 图表题注 · 参考文献比例</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.divider()

    # 文件上传
    st.markdown("#### 📁 上传你的毕业论文终稿")
    st.markdown(
        "<div class='custom-warning'>"
        "<strong>⚠️ 免责声明：</strong>"
        "本工具基于 python-docx 解析 .docx 文件进行格式校验，"
        "结果仅供参考。实际排版标准请以学校最新模板和导师要求为准。"
        "</div>",
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "仅支持 .docx 格式的 Word 文档", type=["docx"], key="thesis_uploader"
    )

    if uploaded_file is not None:
        if st.button(
            "🚀 开始终稿查验", type="primary", use_container_width=True, key="thesis_btn"
        ):
            with st.spinner("🔬 正在调用论文校验引擎，执行全维度深度扫描..."):
                try:
                    with tempfile.NamedTemporaryFile(
                        delete=False, suffix=".docx"
                    ) as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name

                    verifier = ThesisFormatVerifier(tmp_path)
                    result = verifier.run_all()
                    os.remove(tmp_path)

                    st.balloons()
                    _render_thesis_report(result, uploaded_file.name)

                except Exception as e:
                    st.error(f"❌ 校验过程中发生错误，请检查文档是否损坏。错误信息：{e}")


def _render_thesis_report(result: dict, doc_name: str):
    """用 Streamlit 原生组件渲染毕业论文校验报告。"""
    summary = result.get("summary", {})
    total = summary.get("total_checks", 0)
    passed = summary.get("passed", 0)
    failed = summary.get("failed", 0)
    errors = summary.get("errors", 0)
    pass_rate = (passed / total * 100) if total > 0 else 0

    # ── 顶部状态栏 ──
    st.success(f"🎉 校验完成！文件：{doc_name}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("总检查项", total)
    m2.metric("✅ 通过", passed)
    m3.metric("❌ 失败", failed)
    m4.metric("⚠️ 异常", errors)

    st.progress(int(pass_rate), text=f"通过率 {pass_rate:.0f}%")
    st.divider()

    # ── 失败项（红色区域）──
    failed_items = result.get("failed_items", [])
    if failed_items:
        st.markdown(f"### ❌ 待修改项（{len(failed_items)}）")
        for item in failed_items:
            name = item.get("name", "未知项")
            message = item.get("message", "")
            context = item.get("context_text")
            with st.expander(f"❌ {name} — {message[:60]}{'...' if len(message) > 60 else ''}", expanded=False):
                st.markdown(f"**{name}**")
                st.error(message)
                if context:
                    st.code(context, language=None)

    # ── 通过项（绿色区域）──
    passed_items = result.get("passed_items", [])
    if passed_items:
        st.markdown(f"### ✅ 已通过项（{len(passed_items)}）")
        for item in passed_items:
            name = item.get("name", "未知项")
            message = item.get("message", "")
            with st.expander(f"✅ {name}", expanded=False):
                st.success(message)

    # ── 异常/警告（橙色区域）──
    error_list = result.get("errors", [])
    if error_list:
        st.markdown(f"### ⚠️ 异常 / 警告（{len(error_list)}）")
        for err in error_list:
            st.warning(err)


# ═════════════════════════════════════════════════════════════
# 路由分发
# ═════════════════════════════════════════════════════════════
if page == "📝 开题报告格式校验":
    render_proposal_checker()
elif page == "📑 毕业论文终稿查验":
    render_thesis_checker()
