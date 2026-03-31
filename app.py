import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ERP 库存月报自动化整合工具", layout="wide")

st.title("📦 ERP 库存月报自动化整合工具")
st.markdown("""
**功能说明：**
以【本月(3月)报表】为基准，将【上月(2月)报表】中的历史信息（采购订单、入库时间、存放位置等）自动匹配合并。
* ✔️ 自动保留共有的数据，并更新最新库存。
* ✔️ 自动增加本月新增的物料。
* ✔️ 自动剔除本月已不在库（上月有，本月无）的物料。
""")

# 1. 文件上传
col1, col2 = st.columns(2)
with col1:
    feb_file = st.file_uploader("上传上月（2月）ERP库存明细表", type=['xlsx', 'xls', 'csv'])
with col2:
    mar_file = st.file_uploader("上传本月（3月）库存表", type=['xlsx', 'xls', 'csv'])


# 辅助函数：读取文件
def load_data(file):
    if file.name.endswith('.csv'):
        # 尝试不同编码读取
        try:
            return pd.read_csv(file, encoding='utf-8-sig')
        except:
            return pd.read_csv(file, encoding='gbk')
    else:
        return pd.read_excel(file)


if feb_file and mar_file:
    if st.button("🚀 开始自动化整合", type="primary"):
        with st.spinner("正在处理数据，请稍候..."):
            try:
                # 2. 读取数据
                df_feb = load_data(feb_file)
                df_mar = load_data(mar_file)

                # 3. 数据清洗：去除空行或底部“合计”之类的统计行
                # 假设正常的物料编码都是数字或包含数字的有效字符串
                df_feb = df_feb.dropna(subset=['物料编码'])
                df_feb = df_feb[df_feb['物料编码'] != '合计']

                df_mar = df_mar.dropna(subset=['物料'])

                # 4. 统一字段名称 (将3月字段映射为2月的目标字段)
                mar_rename_dict = {
                    '物料': '物料编码',
                    '库存地点': '地点',
                    '基本计量单位': '单位',
                    '非限制使用的库存': '数量',
                    '值未限制': '库存金额'
                }
                df_mar = df_mar.rename(columns=mar_rename_dict)

                # 确保关键列是字符串格式，防止匹配失败
                df_mar['物料编码'] = df_mar['物料编码'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_feb['物料编码'] = df_feb['物料编码'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_mar['批次'] = df_mar['批次'].astype(str)
                df_feb['批次'] = df_feb['批次'].astype(str)

                # 5. 提取2月报表中的历史信息字典 (去除重复项以防报错)
                history_cols = ['物料编码', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注']
                # 如果2月表里有些字段本身就不存在，则忽略
                history_cols = [c for c in history_cols if c in df_feb.columns]
                df_feb_history = df_feb[history_cols].drop_duplicates(subset=['物料编码', '批次'], keep='first')

                # 6. 核心逻辑：左连接 (Left Join)
                # 以3月为主表，去2月里找对应的 订单、入库时间、位置 等信息
                df_merged = pd.merge(df_mar, df_feb_history, on=['物料编码', '批次'], how='left')

                # 7. 整理最终列的顺序，与2月报表一致
                final_columns = [
                    '序号', '工厂', '地点', '物料编码', '物料描述', '单位', '数量',
                    '库存金额', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注'
                ]

                # 如果合出来的表缺了哪些列，用空值补齐
                for col in final_columns:
                    if col not in df_merged.columns:
                        df_merged[col] = None

                df_merged = df_merged[final_columns]

                # 重新生成序号
                df_merged['序号'] = range(1, len(df_merged) + 1)

                st.success("✅ 数据整合成功！")

                # 8. 页面预览
                st.dataframe(df_merged.head(10))

                # 9. 提供下载
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='整合后库存明细')

                st.download_button(
                    label="⬇️ 下载整合后的 Excel 报表",
                    data=output.getvalue(),
                    file_name="整合后_3月ERP库存明细.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"处理过程中出错，请检查表格格式。错误信息: {str(e)}")