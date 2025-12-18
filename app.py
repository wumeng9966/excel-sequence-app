# app_simple.py
import streamlit as st
import pandas as pd
import io
import time
from excel_processor import process_excel_with_sequences

# 设置页面配置
st.set_page_config(
    page_title="Excel序列获取工具",
    page_icon="🧬",
    layout="centered"  # 使用居中布局，更简单
)

# 应用标题
st.title("🧬 Excel序列获取工具")
st.markdown("这是一个简单的工具，用于获取Excel文件中K列和O列的DNA序列。")

# 文件上传
uploaded_file = st.file_uploader("选择Excel文件 (.xlsx)", type=["xlsx"])

# 如果上传了文件
if uploaded_file is not None:
    # 显示基本信息
    st.write(f"**文件:** {uploaded_file.name}")
    st.write(f"**大小:** {uploaded_file.size / 1024:.1f} KB")
    
    # 预览（可选）
    if st.checkbox("预览前5行"):
        try:
            df = pd.read_excel(uploaded_file, nrows=5)
            st.dataframe(df)
        except:
            st.warning("无法预览文件")
    
    # 处理按钮
    if st.button("开始处理序列", type="primary"):
        try:
            # 使用简单的进度指示
            progress_placeholder = st.empty()
            progress_placeholder.text("正在处理，请稍候...")
            
            # 读取文件内容
            file_content = uploaded_file.getvalue()
            
            # 调用处理函数
            start_time = time.time()
            success_count, processed_content = process_excel_with_sequences(file_content)
            end_time = time.time()
            
            if processed_content is not None:
                # 显示结果
                progress_placeholder.empty()
                
                st.success(f"✅ 处理完成!")
                st.write(f"**处理时间:** {end_time - start_time:.1f} 秒")
                st.write(f"**成功获取序列数:** {success_count}")
                
                # 下载按钮
                st.download_button(
                    label="📥 下载结果文件",
                    data=processed_content,
                    file_name=f"processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 简单预览
                if st.checkbox("预览结果前5行"):
                    try:
                        result_df = pd.read_excel(io.BytesIO(processed_content), nrows=5)
                        st.dataframe(result_df)
                    except:
                        st.info("无法预览结果")
            else:
                progress_placeholder.error("❌ 处理失败")
                
        except Exception as e:
            st.error(f"处理过程中发生错误: {str(e)}")
            st.exception(e)
else:
    st.info("👆 请先上传Excel文件")

# 页脚
st.markdown("---")
st.caption("版本 1.0 | 基于Streamlit Cloud部署")
