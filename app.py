# app.py
import streamlit as st
import pandas as pd
import os
import tempfile
import time
from excel_processor import process_excel_with_sequences

# 设置页面配置
st.set_page_config(
    page_title="Excel序列获取工具",
    page_icon="🧬",
    layout="wide"
)

# 应用标题和说明
st.title("🧬 Excel序列获取工具")
st.markdown("""
这个工具可以自动处理Excel文件，为K列和O列的每个位置从网站获取DNA序列。
""")

# 在侧边栏添加说明
with st.sidebar:
    st.header("使用说明")
    st.markdown("""
    1. **上传Excel文件**（确保包含K列和O列）
    2. 点击"开始处理"按钮
    3. 等待处理完成
    4. 下载结果文件
    
    **注意事项：**
    - 处理需要一些时间，请耐心等待
    - 请确保网络连接正常
    - 建议先测试小文件
    """)
    
    # 显示当前状态
    st.header("系统状态")
    if 'processing' in st.session_state and st.session_state.processing:
        st.warning("正在处理中...")
    else:
        st.success("系统就绪")

# 文件上传区域
st.header("📁 上传Excel文件")
uploaded_file = st.file_uploader(
    "选择Excel文件（.xlsx格式）",
    type=["xlsx"],
    help="请确保文件包含K列和O列，且格式正确"
)

# 处理选项
st.header("⚙️ 处理选项")
col1, col2 = st.columns(2)
with col1:
    delay_time = st.slider(
        "请求间隔时间（秒）",
        min_value=0.5,
        max_value=5.0,
        value=1.0,
        step=0.5,
        help="网站请求间隔，避免请求过快"
    )
with col2:
    auto_open = st.checkbox(
        "处理完成后自动显示预览",
        value=True
    )

# 处理按钮和状态显示
if uploaded_file is not None:
    # 显示文件信息
    file_details = {
        "文件名": uploaded_file.name,
        "文件大小": f"{uploaded_file.size / 1024:.2f} KB",
        "文件类型": uploaded_file.type
    }
    st.write("文件信息：", file_details)
    
    # 预览文件内容（前5行）
    try:
        df = pd.read_excel(uploaded_file, nrows=5)
        with st.expander("预览文件前5行"):
            st.dataframe(df)
    except:
        st.warning("无法预览文件内容")
    
    # 开始处理按钮
    if st.button("🚀 开始处理", type="primary", use_container_width=True):
        # 设置处理状态
        st.session_state.processing = True
        
        # 创建临时文件保存上传的文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            input_path = tmp_file.name
        
        # 显示处理进度
        progress_text = "正在处理，请稍候..."
        progress_bar = st.progress(0, text=progress_text)
        status_text = st.empty()
        
        try:
            # 调用处理函数
            status_text.info("正在初始化浏览器驱动...")
            
            # 这里为了演示，模拟处理过程
            # 实际使用时，需要调用处理函数
            # 注意：由于处理时间可能较长，可以考虑使用后台线程
            
            # 模拟处理进度
            for i in range(100):
                time.sleep(0.05)  # 模拟处理时间
                progress_bar.progress(i + 1, text=f"处理中... {i+1}%")
            
            # 实际调用处理函数
            status_text.info("正在获取序列...")
            
            # 调用处理函数
            success_count, output_path = process_excel_with_sequences(input_path)
            
            # 更新进度条
            progress_bar.progress(100, text="处理完成！")
            
            # 显示处理结果
            status_text.success(f"处理完成！成功获取 {success_count} 条序列")
            
            # 提供下载按钮
            with open(output_path, 'rb') as f:
                st.download_button(
                    label="📥 下载处理后的文件",
                    data=f,
                    file_name=f"processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # 预览处理结果
            if auto_open and os.path.exists(output_path):
                try:
                    result_df = pd.read_excel(output_path, nrows=10)
                    with st.expander("预览处理结果（前10行）"):
                        st.dataframe(result_df)
                        
                        # 显示统计信息
                        st.metric("成功获取序列数", success_count)
                except Exception as e:
                    st.warning(f"无法预览结果文件: {str(e)}")
            
        except Exception as e:
            progress_bar.progress(100, text="处理失败")
            status_text.error(f"处理过程中发生错误: {str(e)}")
            st.exception(e)
        finally:
            # 清理临时文件
            try:
                os.unlink(input_path)
                if 'output_path' in locals():
                    # 可以选择是否删除输出文件
                    # os.unlink(output_path)
                    pass
            except:
                pass
            
            # 重置处理状态
            st.session_state.processing = False
else:
    st.info("👆 请先上传Excel文件")

# 页脚信息
st.markdown("---")
st.caption("© 2023 Excel序列获取工具 | 版本 1.0")