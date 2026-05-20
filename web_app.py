"""
培正开题报告格式校验助手 - 网页 Web 版 (Streamlit 驱动)
用法：本地运行 streamlit run web_app.py
"""
import streamlit as st
import os
import tempfile
from Format_verify_tool import analyze_proposal, HTMLReporter

# 页面基础配置（设置网页标签标题和图标）
st.set_page_config(
    page_title="培正学院 - 开题报告格式校验助手",
    page_icon="🎓",
    layout="centered"
)

# ── 1. 顶部视觉区 ──
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>🎓 培正学院开题报告格式校验助手</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6B7280; font-size: 14px;'>毕业季神器 · 云端运行 · 保护隐私</p>", unsafe_allow_html=True)
st.divider()

# ── 2. 左侧/上方 自定义配置面板 ──
st.markdown("#### ⚙️ 高级自定义参数配置(可选)")

col1, col2 = st.columns(2)
with col1:
    line_spacing = st.number_input("预期行距", min_value=1.0, max_value=3.0, value=1.5, step=0.05)
    min_refs = st.number_input("文献数量下限 (篇)", min_value=1, max_value=50, value=10, step=1)
with col2:
    min_words = st.number_input("综述字数下限", min_value=500, max_value=10000, value=1800, step=100)
    max_words = st.number_input("综述字数上限", min_value=500, max_value=20000, value=2200, step=100)

# 弹性规则开关
st.markdown("##### 🛠️ 规则开关控制")
c1, c2, c3 = st.columns(3)
with c1:
    check_tutor_space = st.checkbox("导师职称空格规范", value=True)
with c2:
    check_indent = st.checkbox("参考文献悬挂缩进", value=True)
with c3:
    check_timeline = st.checkbox("进度安排时间线顺叙", value=True)

st.divider()

# ── 3. 文件上传核心交互区 ──
st.markdown("#### 📁 上传你的开题报告")
uploaded_file = st.file_uploader("仅支持 .docx 格式的 Word 文档", type=["docx"])

if uploaded_file is not None:
    # 构建传给大脑的配置字典
    custom_config = {
        'line_spacing': line_spacing,
        'min_refs': min_refs,
        'min_words': min_words,
        'max_words': max_words,
        'check_tutor_space': check_tutor_space,
        'check_indent': check_indent,
        'check_timeline': check_timeline
    }
    
    # 按钮激活
    st.warning("⚠️ 免责声明：本工具格式校验结果仅供参考。由于文档结构可能存在原生性错乱，实际通过标准请以导师最终意见为准，建议生成报告后人工复查一遍。")

    if st.button("🚀 开始一键格式校验", type="primary", use_container_width=True):
        with st.spinner("正在深度解析文档并交叉校对中，请稍候..."):
            try:
                # 因为 Streamlit 上传的是内存文件对象，我们需要把它写入临时文件中供 python-docx 读取
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name
                
                # 1. 运行核心算法大脑
                result = analyze_proposal(tmp_path, config=custom_config)
                
                # 2. 生成 HTML 诊断报告
                doc_name = uploaded_file.name
                reporter = HTMLReporter(result, doc_name=doc_name)
                # 修改：直接让它返回或重新读取生成的 report.html 内容
                report_path = reporter.generate("web_report.html")
                
                with open(report_path, "r", encoding="utf-8") as f:
                    html_content = f.read()
                
                # 清理临时文件
                os.remove(tmp_path)
                os.remove(report_path)
                
                # 3. 渲染结果
                st.balloons() # 撒花特效
                st.success("🎉 校验完成！诊断报告已在下方生成。")
                
                # 直接将你写的那个漂亮 HTML 报告内嵌进网页中展示！
                st.components.v1.html(html_content, height=900, scrolling=True)
                
            except Exception as e:
                st.error(f"❌ 校验过程中发生错误，请检查文档是否损坏。错误信息：{e}")

# ── 4. 专属页脚署名 ──
st.markdown(
    "<br><hr><p style='text-align: center; color: #9CA3AF; font-size: 14px;'>"
    "Developed with ❤️ by 林格 | 谨以此工具献给被格式折磨的同学们<br><br>"
    "🌟 <a href='https://github.com/linjunhao024-byte/The_All_Web' target='_blank' style='color: #4F46E5; text-decoration: none; font-weight: bold;'>本项目源代码已在 GitHub 开源，欢迎访问与交流！</a>"
    "</p>", 
    unsafe_allow_html=True
)