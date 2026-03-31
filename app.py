import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ERP 库存月报自动化整合工具", layout="wide")

st.title("📦 ERP 库存月报自动化整合工具")
st.markdown("""
**功能说明：**
以【最新报表】为基准，将【历史报表】中的附加信息（采购订单、入库时间、存放位置等）自动匹配合并。
* ✔️ 自动保留共有的数据，并同步历史信息。
* ✔️ 自动增加最新表中的新增物料。
* ✔️ 自动剔除已不在最新表中的物料。
* 🤖 **智能防错**：跳过大标题、清理隐形空格、自动完美列宽、**自动去除多余的00:00:00时间尾巴**。
""")

# 1. 文件上传
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 上传【历史】库存表 (如上月)", type=['xlsx', 'xls', 'csv'])
with col2:
    new_file = st.file_uploader("🆕 上传【最新】库存表 (如本月)", type=['xlsx', 'xls', 'csv'])

# 辅助函数：读取文件并智能定位真实表头
def load_data(file):
    if file.name.endswith('.csv'):
        try:
            df = pd.read_csv(file, encoding='utf-8-sig')
        except:
            df = pd.read_csv(file, encoding='gbk')
    else:
        df = pd.read_excel(file)
        
    if '物料编码' not in df.columns and '物料' not in df.columns:
        for i in range(min(5, len(df))):
            row_values = [str(val).strip() for val in df.iloc[i].values]
            if '物料编码' in row_values or '物料' in row_values:
                df.columns = row_values
                df = df.iloc[i+1:].reset_index(drop=True)
                break
                
    df.columns = [str(col).strip() for col in df.columns]
    return df

if old_file and new_file:
    if st.button("🚀 开始自动化整合", type="primary"):
        with st.spinner("正在处理数据，请稍候..."):
            try:
                # 2. 读取数据
                df_old = load_data(old_file)
                df_new = load_data(new_file)

                # 3. 数据清洗：去除空行或底部“合计”行
                df_old = df_old.dropna(subset=['物料编码'])
                df_old = df_old[df_old['物料编码'].astype(str).str.strip() != '合计']
                
                df_new = df_new.dropna(subset=['物料'])
                
                # 4. 统一字段名称
                new_rename_dict = {
                    '物料': '物料编码',
                    '库存地点': '地点',
                    '基本计量单位': '单位',
                    '非限制使用的库存': '数量',
                    '值未限制': '库存金额'
                }
                df_new = df_new.rename(columns=new_rename_dict)

                # --- 核心防御性加固：彻底清理 NaN 和隐形空格 ---
                df_new['物料编码'] = df_new['物料编码'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '').str.strip()
                df_old['物料编码'] = df_old['物料编码'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '').str.strip()
                df_new['批次'] = df_new['批次'].fillna('').astype(str).replace('nan', '').str.strip()
                df_old['批次'] = df_old['批次'].fillna('').astype(str).replace('nan', '').str.strip()

                # 5. 提取历史信息
                history_cols = ['物料编码', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注']
                history_cols = [c for c in history_cols if c in df_old.columns]
                df_old_history = df_old[history_cols].drop_duplicates(subset=['物料编码', '批次'], keep='first')

                # 6. 左连接合并
                df_merged = pd.merge(df_new, df_old_history, on=['物料编码', '批次'], how='left')

                # 7. 整理列顺序
                final_columns = [
                    '序号', '工厂', '地点', '物料编码', '物料描述', '单位', '数量', 
                    '库存金额', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注'
                ]
                
                for col in final_columns:
                    if col not in df_merged.columns:
                        df_merged[col] = None
                        
                df_merged = df_merged[final_columns]
                df_merged['序号'] = range(1, len(df_merged) + 1)
                
                # --- 新增：彻底消灭时间尾巴 00:00:00 ---
                if '入库时间' in df_merged.columns:
                    # 先转为纯文本，然后用正则替换掉所有 00:00:00 及其前面的空格
                    df_merged['入库时间'] = df_merged['入库时间'].astype(str)
                    df_merged['入库时间'] = df_merged['入库时间'].str.replace(r'\s*00:00:00$', '', regex=True)
                    # 将空值（nan/None/NaT）替换为空白
                    df_merged['入库时间'] = df_merged['入库时间'].replace(['nan', 'None', 'NaT'], '')
                # ------------------------------------

                st.success("✅ 数据整合成功！")
                st.dataframe(df_merged.head(10))

                # 8. 导出并自动调整列宽
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='整合后库存明细')
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['整合后库存明细']
                    
                    # 遍历所有列，动态设置列宽
                    for i, col in enumerate(df_merged.columns):
                        col_data = df_merged[col].astype(str).replace('nan', '').replace('None', '')
                        max_len = max(col_data.map(len).max(), len(str(col)))
                        
                        if col in ['物料描述', '备注']:
                            set_len = min(max_len * 1.5, 60)
                        elif col in ['物料编码', '批次', '入库时间']:
                            set_len = max(max_len + 4, 15) # 时间修剪后，15的宽度绰绰有余
                        else:
                            set_len = max_len + 4
                            
                        worksheet.set_column(i, i, set_len)
                
                # 9. 提供下载
                original_name = new_file.name.rsplit('.', 1)[0]
                download_file_name = f"整合后_{original_name}.xlsx"
                
                st.download_button(
                    label=f"⬇️ 下载 {download_file_name}",
                    data=output.getvalue(),
                    file_name=download_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"处理过程中出错，请检查表格格式。错误信息: {str(e)}")
