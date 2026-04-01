import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="ERP 库存月报自动化整合工具", layout="wide")

# --- 1. 注入 CSS 修改上传按钮文字 ---
st.markdown("""
    <style>
    section[data-testid="stFileUploadDropzone"] button {
        visibility: hidden;
    }
    section[data-testid="stFileUploadDropzone"] button:before {
        content: "上传文件";
        visibility: visible;
        display: block;
        background-color: #FF4B4B;
        color: white;
        padding: 8px 16px;
        border-radius: 8px;
    }
    section[data-testid="stFileUploadDropzone"] span {
        display: none;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 定义数据读取函数 ---
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

# --- 3. 业务 UI 界面 ---
st.title("📦 ERP 库存月报自动化整合工具 V1.1") # 增加了版本号便于确认更新

st.info("""
**💡 操作指引：**
1. **上传历史表**：上传上月库存表（用于提取历史入库时间、存放位置等）。
2. **上传最新表**：上传本月导出的最新库存表（作为整合基准）。
3. **一键生成**：点击下方按钮，系统将自动完成匹配并生成带标题的报表。
""")

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 第一步：上传【历史】库存表", type=['xlsx', 'xls', 'csv'])
with col2:
    new_file = st.file_uploader("🆕 第二步：上传【最新】库存表", type=['xlsx', 'xls', 'csv'])

if old_file and new_file:
    st.write("---")
    if st.button("🚀 开始自动化整合", type="primary", use_container_width=True):
        with st.spinner("系统正在处理数据并排版..."):
            try:
                df_old = load_data(old_file)
                df_new = load_data(new_file)

                df_old = df_old.dropna(subset=['物料编码'])
                df_old = df_old[df_old['物料编码'].astype(str).str.strip() != '合计']
                df_new = df_new.dropna(subset=['物料'])
                
                new_rename_dict = {
                    '物料': '物料编码',
                    '库存地点': '地点',
                    '基本计量单位': '单位',
                    '非限制使用的库存': '数量',
                    '值未限制': '库存金额'
                }
                df_new = df_new.rename(columns=new_rename_dict)

                for df_tmp in [df_new, df_old]:
                    col_target = '物料编码' if '物料编码' in df_tmp.columns else '物料'
                    df_tmp[col_target] = df_tmp[col_target].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '').str.strip()
                    if '批次' in df_tmp.columns:
                        df_tmp['批次'] = df_tmp['批次'].fillna('').astype(str).replace('nan', '').str.strip()

                history_cols = ['物料编码', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注']
                history_cols = [c for c in history_cols if c in df_old.columns]
                df_old_history = df_old[history_cols].drop_duplicates(subset=['物料编码', '批次'], keep='first')

                df_merged = pd.merge(df_new, df_old_history, on=['物料编码', '批次'], how='left')

                final_columns = [
                    '序号', '工厂', '地点', '物料编码', '物料描述', '单位', '数量', 
                    '库存金额', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注'
                ]
                for col in final_columns:
                    if col not in df_merged.columns:
                        df_merged[col] = None
                df_merged = df_merged[final_columns]
                df_merged['序号'] = range(1, len(df_merged) + 1)
                
                if '入库时间' in df_merged.columns:
                    def to_datetime_obj(d):
                        if pd.isna(d) or str(d).strip() in ['', 'nan', 'None', 'NaT']:
                            return pd.NaT 
                        d_str = str(d).strip().split(' ')[0]
                        d_str = d_str.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-')
                        d_str = d_str.rstrip('-')
                        if len(d_str.split('-')) == 2: d_str += '-01'
                        try: return pd.to_datetime(d_str)
                        except: return str(d).strip()
                    df_merged['入库时间'] = df_merged['入库时间'].apply(to_datetime_obj)

                year_match = re.search(r'(20\d{2})', new_file.name)
                year_str = year_match.group(1) if year_match else "2026"
                month_match = re.search(r'(\d+)月', new_file.name)
                month_str = month_match.group(1) if month_match else "X"
                report_title = f"天津液化{year_str}年{month_str}月ERP库存明细表"

                st.success("✅ 整合完成！")
                st.dataframe(df_merged.head(10))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='yyyy"年"m"月"d"日"') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='库存明细', startrow=1)
                    workbook  = writer.book
                    worksheet = writer.sheets['库存明细']
                    title_fmt = workbook.add_format({'bold':True, 'font_size':16, 'align':'center', 'valign':'vcenter'})
                    worksheet.merge_range(0, 0, 0, len(df_merged.columns)-1, report_title, title_fmt)
                    worksheet.set_row(0, 30)
                    
                    # --- 终极加固：计算列宽 ---
                    for i, col in enumerate(df_merged.columns):
                        # 1. 强制将该列所有元素转为字符串（包括 NaN）
                        # 2. 用 lambda 确保 len() 只作用于字符串对象
                        # 3. 过滤掉无意义的 'nan', 'None', 'NaT' 后再算长度
                        s_list = df_merged[col].astype(str).replace(['nan', 'None', 'NaT', '<NA>'], '')
                        
                        # 找出内容的最长长度（如果没有内容则为 0）
                        if len(s_list) > 0:
                            # 关键修复点：使用 lambda 兜底转换
                            max_content_len = s_list.apply(lambda x: len(str(x))).max()
                        else:
                            max_content_len = 0
                        
                        # 比较标题长度
                        max_tick_len = max(max_content_len, len(str(col)))
                        
                        # 设置不同字段的宽度权重
                        if col in ['物料描述', '备注']:
                            width = min(max_tick_len * 1.5, 60)
                        elif col in ['物料编码', '批次', '入库时间']:
                            width = max(max_tick_len + 4, 18)
                        else:
                            width = max_tick_len + 4
                            
                        worksheet.set_column(i, i, width)

                st.download_button(
                    label=f"⬇️ 下载整合报表",
                    data=output.getvalue(),
                    file_name=f"整合后_{new_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            except Exception as e:
                # 即使报错也把具体的报错行显示出来，方便定位
                import traceback
                st.error(f"处理过程中出现错误：{e}")
                st.code(traceback.format_exc())
