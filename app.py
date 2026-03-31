import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="ERP 库存月报自动化整合工具", layout="wide")

# --- 注入 CSS 强行修改按钮文字 ---
st.markdown("""
    <style>
    /* 定位到上传组件中的按钮文本 */
    section[data-testid="stFileUploadDropzone"] button {
        visibility: hidden;
    }
    section[data-testid="stFileUploadDropzone"] button:before {
        content: "上传文件";
        visibility: visible;
        display: block;
        background-color: #FF4B4B; /* 这里可以自定义颜色，默认是红 */
        color: white;
        padding: 8px 16px;
        border-radius: 8px;
    }
    /* 隐藏底部的 "Browse files" 默认提示 */
    section[data-testid="stFileUploadDropzone"] span {
        display: none;
    }
    </style>
    """, unsafe_allow_html=True)

# ================= 优化后的极简业务 UI =================
st.title("📦 ERP 库存月报自动化整合工具")

st.info("""
**💡 操作指引：**
1. **上传历史表**：上传上月库存表（用于提取原有的入库时间、存放位置、采购订单等辅助信息）。
2. **上传最新表**：上传本月系统导出的最新库存表（系统将自动以该表为基准进行匹配与更新）。
3. **一键生成**：点击“开始自动化整合”，即可下载格式统一、排版完成的最终报表。
""")

# 1. 文件上传区
col1, col2 = st.columns(2)
with col1:
    # 注意：这里的 label 我们依然写中文，CSS 会处理按钮内的文字
    old_file = st.file_uploader("📂 第一步：上传【历史】库存表", type=['xlsx', 'xls', 'csv'])
with col2:
    new_file = st.file_uploader("🆕 第二步：上传【最新】库存表", type=['xlsx', 'xls', 'csv'])

# ... (后面的数据处理逻辑 df_old, df_new 等保持不变)

if old_file and new_file:
    st.write("---")
    if st.button("🚀 开始自动化整合", type="primary", use_container_width=True):
        # ... (之前确认无误的所有处理逻辑)
        with st.spinner("系统正在高速处理并排版数据，请稍候..."):
            try:
                # 2. 读取数据 (这里沿用之前的逻辑)
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
                st.markdown("###### 👁️ 整合结果预览 (前10条)")
                st.dataframe(df_merged.head(10))

                # 智能提取月份
                year_match = re.search(r'(20\d{2})年', new_file.name)
                year_str = year_match.group(1) if year_match else "2026"
                month_match = re.search(r'(\d+)月', new_file.name)
                month_str = month_match.group(1) if month_match else "X"
                report_title = f"天津液化{year_str}年{month_str}月ERP库存明细表"

                # 8. 导出
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='yyyy"年"m"月"d"日"') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='整合后库存明细', startrow=1)
                    workbook  = writer.book
                    worksheet = writer.sheets['整合后库存明细']
                    title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
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
                
                # 9. 下载
                original_name = new_file.name.rsplit('.', 1)[0]
                download_file_name = f"整合后_{original_name}.xlsx"
                st.write("---")
                st.download_button(
                    label=f"⬇️ 下载最终报表：{download_file_name}",
                    data=output.getvalue(),
                    file_name=download_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            except Exception as e:
                st.error(f"处理出错：{str(e)}")
