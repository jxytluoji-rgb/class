import streamlit as st
import pandas as pd
import os
import tempfile
import base64
from datetime import datetime

from utils import (
    convert_classroom_schedule,
    read_source_file_classroom,
    generate_classroom_sign
)

st.set_page_config(
    page_title="课表生成器",
    page_icon="📚",
    layout="wide"
)

def get_download_link(file_path, filename):
    with open(file_path, 'rb') as f:
        file_data = f.read()
    b64 = base64.b64encode(file_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">⬇️ 下载 {filename}</a>'
    return href

@st.cache_data
def load_placeholder_data():
    return pd.DataFrame({
        '功能': ['教室课表', '教室班牌'],
        '说明': ['适合教师和管理员查看的详细课表', '适合在教室门口显示的简洁课表']
    })

def main():
    st.title("📚 课表生成器")
    st.markdown("---")

    with st.sidebar:
        st.header("⚙️ 设置")
        convert_type = st.selectbox(
            "转换类型",
            ["教室课表", "教室班牌"],
            help="选择要生成的课表类型"
        )

        target_week = st.number_input(
            "选择周数",
            min_value=1,
            max_value=20,
            value=1,
            help="输入要生成的周数"
        )

        st.markdown("---")
        st.markdown("""
        ### 📖 使用说明
        1. 上传Excel源文件
        2. 选择转换类型
        3. 设置周数
        4. 点击生成按钮
        5. 下载生成的课表
        """)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📤 上传课表源文件")
        uploaded_file = st.file_uploader(
            "请选择Excel文件",
            type=['xlsx', 'xls'],
            help="支持.xlsx和.xls格式的Excel文件"
        )

    with col2:
        if uploaded_file is not None:
            st.success(f"已选择: {uploaded_file.name}")
            st.write(f"大小: {uploaded_file.size / 1024:.1f} KB")

    if uploaded_file is not None:
        if st.button("🚀 生成课表", type="primary", use_container_width=True):
            with st.spinner("正在生成课表，请稍候..."):
                try:
                    tfile = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                    tfile.write(uploaded_file.getbuffer())
                    tfile.close()

                    output_filename = f"第{target_week}周{'课表' if convert_type == '教室课表' else '班牌'}.xlsx"
                    output_path = os.path.join(tempfile.gettempdir(), output_filename)

                    if convert_type == "教室课表":
                        record_count = convert_classroom_schedule(
                            tfile.name,
                            output_path,
                            target_week=target_week
                        )
                    else:
                        course_data = read_source_file_classroom(tfile.name, target_week)
                        record_count = generate_classroom_sign(course_data, output_path)

                    os.unlink(tfile.name)

                    st.success(f"✅ 转换完成！共生成 {record_count} 条记录")

                    st.markdown("### 📥 下载课表")
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        st.markdown(
                            get_download_link(output_path, output_filename),
                            unsafe_allow_html=True
                        )

                    st.info(f"📁 文件路径: {output_path}")

                except Exception as e:
                    st.error(f"❌ 转换失败: {str(e)}")

    st.markdown("---")

    with st.expander("📋 功能特点", expanded=False):
        st.markdown("""
        - ✅ **教室课表**: 适合教师和管理员查看的详细课表
        - ✅ **教室班牌**: 适合在教室门口显示的简洁课表
        - ✅ 自动处理单双周课程
        - ✅ 数据自动居中显示
        - ✅ 空白字段自动替换为"空闲"
        - ✅ 教室班牌包含设备ID映射
        - ✅ 支持自定义周数选择（1-20周）
        """)

    st.markdown("---")
    st.caption("课表生成器 v2.0 | 基于 Streamlit 构建")

if __name__ == "__main__":
    main()
