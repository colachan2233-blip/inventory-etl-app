import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="ERP 库存月报自动化整合工具", layout="wide")

# ================= 优化后的极简业务 UI =================
st.title("📦 ERP 库存月报自动化整合工具")

# 使用 info 提示框让指引更醒目、专业
st.info("""
**💡 操作指引：**
1. **上传历史表**：上传上月库存表（用于提取原有的入库时间、存放位置、采购订单等辅助信息）。
2. **上传最新表**：上传本月系统导出的最新库存表（系统将自动以该表为基准进行匹配与更新）。
3. **一键生成**：点击“开始自动化整合”，即可下载格式统一、排版完成的最终报表。
""")
# =======================================================

# 1. 文件上传区
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 第一步：上传【历史】库存表", type=['xlsx', 'xls', 'csv'])
with col2:
    new_file = st.file_uploader("🆕 第二步：上传【最新】库存表", type=['xlsx', 'xls', 'csv'])

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

# 主逻辑
if old_file and new_file:
    # 按钮增加了一点间距，并且文案更具行动力
    st.write("---")
    if st.button("🚀 开始自动化整合", type="primary", use_container_width=True):
        with st.spinner("系统正在高速处理并排版数据，请稍候..."):
            try:
                # 2. 读取数据
                df_old = load_data(old_file)
                df_new = load_data(new_file)

                # 3. 数据清洗
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

                # 核心防御：清理 NaN 和隐形空格
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
                
                # 写入真实的底层 DateTime 对象
                if '入库时间' in df_merged.columns:
                    def to_real_excel_date(d):
                        if pd.isna(d) or str(d).strip() in ['', 'nan', 'None', 'NaT']:
                            return pd.NaT 
                            
                        d_str = str(d).strip().split(' ')[0]
                        d_str = d_str.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-')
                        d_str = d_str.rstrip('-')
                        
                        if len(d_str.split('-')) == 2:
                            d_str += '-01'
                            
                        try:
                            return pd.to_datetime(d_str)
                        except:
                            return str(d).strip()
                            
                    df_merged['入库时间'] = df_merged['入库时间'].apply(to_real_excel_date)

                st.success("✅ 报表已生成！请点击下方按钮下载。")
                
                # 界面优化：预览表格前加上一个小标题
                st.markdown("###### 👁️ 整合结果预览 (前10条)")
                st.dataframe(df_merged.head(10))

                # 智能提取年份和月份生成大标题
                year_match = re.search(r'(20\d{2})年', new_file.name)
                year_str = year_match.group(1) if year_match else "2026"
                
                month_match = re.search(r'(\d+)月', new_file.name)
                month_str = month_match.group(1) if month_match else "X"
                
                report_title = f"天津液化{year_str}年{month_str}月ERP库存明细表"

                # 8. 导出：注入 Excel 自定义时间皮肤，并添加大标题
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='yyyy"年"m"月"d"日"') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='整合后库存明细', startrow=1)
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['整合后库存明细']
                    
                    title_format = workbook.add_format({
                        'bold': True,
                        'font_size': 16,
                        'align': 'center',
                        'valign': 'vcenter'
                    })
                    
                    worksheet.merge_range(0, 0, 0, len(df_merged.columns) - 1, report_title, title_format)
                    worksheet.set_row(0, 30) 
                    
                    for i, col in enumerate(df_merged.columns):
                        col_data = df_merged[col].astype(str).replace(['nan', 'None', 'NaT'], '')
                        max_len = max(col_data.map(len).max(), len(str(col)))
                        
                        if col in ['物料描述', '备注']:
                            set_len = min(max_len * 1.5, 60)
                        elif col in ['物料编码', '批次', '入库时间']:
                            set_len = max(max_len + 4, 18)
                        else:
                            set_len = max_len + 4
                            
                        worksheet.set_column(i, i, set_len)
                
                # 9. 提供下载
                original_name = new_file.name.rsplit('.', 1)[0]
                download_file_name = f"整合后_{original_name}.xlsx"
                
                st.write("---") # 加一条分割线让下载区域更独立
                st.download_button(
                    label=f"⬇️ 下载最终报表：{download_file_name}",
                    data=output.getvalue(),
                    file_name=download_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary" # 让下载按钮也变成醒目的主色调
                )

            except Exception as e:
                st.error(f"处理过程中出错，请检查表格格式是否规范。系统提示: {str(e)}")
