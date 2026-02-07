import pandas as pd
import json
import streamlit as st
from io import BytesIO

def convert_txt_to_excel(txt_content):
    """
    将TXT文件中的JSON数据转换为Excel格式
    参数: txt_content - 读取的TXT文件内容（字符串）
    返回: Excel文件的BytesIO对象，可直接用于下载
    """
    try:
        # 解析JSON数据
        data = json.loads(txt_content)
        # 转换为DataFrame
        df = pd.DataFrame(data['data'])
        
        # 处理rankingDataList数据
        result = pd.DataFrame()
        for i in df['rankingDataList']:
            r = i
            for j in i:
                r[j] = [r[j]]
            result = pd.concat([result, pd.DataFrame(r, index=[0])], ignore_index=True)
        
        # 将DataFrame写入BytesIO（内存中），避免本地文件存储
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result.to_excel(writer, index=None, sheet_name='转换结果')
        # 重置指针到开头，确保下载时能读取完整内容
        output.seek(0)
        
        return output, None  # 返回Excel对象和错误信息（无错误则为None）
    
    except json.JSONDecodeError:
        return None, "文件内容不是有效的JSON格式，请检查TXT文件！"
    except KeyError as e:
        return None, f"JSON数据格式错误，缺少关键字段：{e}！"
    except Exception as e:
        return None, f"转换失败：{str(e)}！"

# 网页应用主体
st.title("TXT(JSON)转Excel工具")
st.markdown("### 上传包含JSON格式的TXT文件，自动转换为Excel并下载")

# 文件上传组件
uploaded_file = st.file_uploader("选择TXT文件", type=["txt"])

if uploaded_file is not None:
    # 读取上传的文件内容
    try:
        txt_content = uploaded_file.getvalue().decode("utf-8")
        st.success("文件上传成功！开始转换...")
        
        # 调用转换函数
        excel_file, error_msg = convert_txt_to_excel(txt_content)
        
        if excel_file:
            # 生成下载按钮
            st.success("转换完成！点击下方按钮下载Excel文件")
            st.download_button(
                label="下载Excel文件",
                data=excel_file,
                file_name=f"{uploaded_file.name.replace('.txt', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # 显示错误信息
            st.error(f"转换失败：{error_msg}")
    
    except UnicodeDecodeError:
        st.error("文件编码错误！请确保TXT文件是UTF-8编码")
    except Exception as e:
        st.error(f"文件读取失败：{str(e)}")

# 侧边栏说明
st.sidebar.markdown("### 使用说明")
st.sidebar.markdown("1. 准备包含JSON格式的TXT文件（需有data/rankingDataList层级）")
st.sidebar.markdown("2. 上传文件后等待转换")
st.sidebar.markdown("3. 转换成功后点击下载按钮获取Excel文件")