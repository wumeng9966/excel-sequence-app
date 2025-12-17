# excel_processor.py
import os
import time
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from openpyxl import load_workbook
# 添加 webdriver-manager 来自动管理驱动
from webdriver_manager.microsoft import EdgeChromiumDriverManager

def setup_driver():
    """配置Edge浏览器驱动，使用webdriver-manager自动管理驱动"""
    options = Options()
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    # 添加无头模式选项（在服务器上运行时不需要图形界面）
    options.add_argument("--headless")
    
    # 使用webdriver-manager自动下载和管理Edge驱动
    service = Service(EdgeChromiumDriverManager().install())
    
    driver = webdriver.Edge(service=service, options=options)
    driver.implicitly_wait(10)
    return driver

def get_sequence_from_website(driver, input_value):
    """
    从网站获取序列
    
    Args:
        driver: 浏览器驱动
        input_value: 输入值，格式如"chr4B:425000640-425000640"
    
    Returns:
        str: 获取的序列，如果失败则返回空字符串
    """
    try:
        # 打开网站
        driver.get('http://202.194.139.32/getfasta/index.html')
        
        # 选择数据库
        db_select = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, 'database')))
        db_select.find_element(By.CSS_SELECTOR, 'option[value="Chinese_Spring1.0.genome"]').click()
        
        # 输入ID
        textarea = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.NAME, 'ID')))
        textarea.clear()
        textarea.send_keys(input_value)
        
        # 提交查询
        submit_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"]')))
        submit_btn.click()
        
        # 获取序列
        seq_div = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, 'seq')))
        
        # 解析序列内容
        seq_text = seq_div.text
        
        # 提取序列部分（去掉第一行的标题）
        if seq_text:
            lines = seq_text.split('\n')
            if len(lines) > 1:
                # 合并所有序列行（去掉第一行的标题）
                sequence = ''.join(lines[1:])
                return sequence
            else:
                print(f"获取的序列格式异常: {seq_text}")
                return ""
        else:
            print("获取的序列为空")
            return ""
            
    except Exception as e:
        print(f"从网站获取序列失败: {str(e)}")
        traceback.print_exc()
        return ""

def process_excel_with_sequences(input_file_path, output_file_path=None):
    """
    处理Excel文件，为每一行的K列和O列获取序列
    
    Args:
        input_file_path: 输入Excel文件路径
        output_file_path: 输出Excel文件路径，如果为None则自动生成
    
    Returns:
        tuple: (成功处理的行数, 输出文件路径)
    """
    # 如果未指定输出文件路径，自动生成
    if output_file_path is None:
        base_name = os.path.splitext(input_file_path)[0]
        output_file_path = f"{base_name}_序列获取后.xlsx"
    
    # 读取输入文件，使用data_only=True来获取公式计算结果
    wb = load_workbook(input_file_path, data_only=True)
    ws = wb.active
    
    # 设置浏览器驱动
    driver = setup_driver()
    
    processed_count = 0
    total_rows = ws.max_row
    success_count = 0
    
    try:
        # 从第1行开始处理（假设第1行有数据）
        for row_idx in range(1, total_rows + 1):
            print(f"处理第 {row_idx}/{total_rows} 行...")
            
            # ========== 处理K列 ==========
            k_value = ws.cell(row=row_idx, column=11).value  # K列是第11列
            if k_value:
                try:
                    # 转换输入值为字符串
                    input_str = str(k_value).strip()
                    
                    # 检查输入格式是否正确（应该包含冒号和破折号）
                    if ":" in input_str and "-" in input_str:
                        sequence = get_sequence_from_website(driver, input_str)
                        
                        if sequence:
                            ws.cell(row=row_idx, column=12, value=sequence)  # L列是第12列
                            success_count += 1
                        else:
                            ws.cell(row=row_idx, column=12, value="获取失败")
                    else:
                        ws.cell(row=row_idx, column=12, value="格式错误")
                except Exception as e:
                    ws.cell(row=row_idx, column=12, value="处理出错")
            else:
                ws.cell(row=row_idx, column=12, value="空值")
            
            # 短暂等待，避免对网站造成过大压力
            time.sleep(1)
            
            # ========== 处理O列 ==========
            o_value = ws.cell(row=row_idx, column=15).value  # O列是第15列
            if o_value:
                try:
                    # 转换输入值为字符串
                    input_str = str(o_value).strip()
                    
                    # 检查输入格式是否正确（应该包含冒号和破折号）
                    if ":" in input_str and "-" in input_str:
                        sequence = get_sequence_from_website(driver, input_str)
                        
                        if sequence:
                            ws.cell(row=row_idx, column=16, value=sequence)  # P列是第16列
                            success_count += 1
                        else:
                            ws.cell(row=row_idx, column=16, value="获取失败")
                    else:
                        ws.cell(row=row_idx, column=16, value="格式错误")
                except Exception as e:
                    ws.cell(row=row_idx, column=16, value="处理出错")
            else:
                ws.cell(row=row_idx, column=16, value="空值")
            
            # 短暂等待
            time.sleep(1)
            
            processed_count += 1
            
            # 每处理10行保存一次，防止数据丢失
            if processed_count % 10 == 0:
                wb.save(output_file_path)
    
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        traceback.print_exc()
    
    finally:
        # 最终保存和清理
        wb.save(output_file_path)
        driver.quit()
        
        return success_count, output_file_path

# 如果直接运行这个文件，可以测试功能
if __name__ == '__main__':
    # 测试代码
    input_file = '429.xlsx'
    if os.path.exists(input_file):
        success_count, output_file = process_excel_with_sequences(input_file)
        print(f"处理完成！成功获取 {success_count} 条序列")
        print(f"结果已保存到: {output_file}")
    else:
        print(f"找不到输入文件: {input_file}")