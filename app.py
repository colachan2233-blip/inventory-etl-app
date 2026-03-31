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
* 🤖 **智能防错**：自动识别并跳过表格顶部的无效大标题。
""")

# 1. 文件上传 (文案改为通用名称)
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("📂 上传【历史】库存表 (如上月)", type=['xlsx', 'xls', 'csv'])
with col2:
    new_file = st.file_uploader("🆕 上传【最新】库存表 (如本月)", type=['xlsx', 'xls', 'csv'])


# 辅助函数：读取文件并智能定位真实表头
def load_data(file):
    # 读取文件
    if file.name.endswith('.csv'):
        try:
            df = pd.read_csv(file, encoding='utf-8-sig')
        except:
            df = pd.read_csv(file, encoding='gbk')
    else:
        df = pd.read_excel(file)

    # 智能寻找真实表头：如果当前列名不对，往下找前5行，看哪一行包含'物料编码'或'物料'
    if '物料编码' not in df.columns and '物料' not in df.columns:
        for i in range(min(5, len(df))):
            row_values = [str(val).strip() for val in df.iloc[i].values]
            if '物料编码' in row_values or '物料' in row_values:
                df.columns = row_values  # 把这一行设为表头
                df = df.iloc[i + 1:].reset_index(drop=True)  # 删掉表头及以上的无效行
                break

    # 兜底清理：去除列名两端的空格，防止因为多敲了空格导致报错
    df.columns = [str(col).strip() for col in df.columns]
    return df


if old_file and new_file:
    if st.button("🚀 开始自动化整合", type="primary"):
        with st.spinner("正在处理数据，请稍候..."):
            try:
                # 2. 读取数据 (变量名也改为 old 和 new)
                df_old = load_data(old_file)
                df_new = load_data(new_file)

                # 3. 数据清洗：去除空行或底部“合计”之类的统计行
                df_old = df_old.dropna(subset=['物料编码'])
                df_old = df_old[df_old['物料编码'].astype(str).str.strip() != '合计']

                df_new = df_new.dropna(subset=['物料'])

                # 4. 统一字段名称 (将新表字段映射为目标字段)
                new_rename_dict = {
                    '物料': '物料编码',
                    '库存地点': '地点',
                    '基本计量单位': '单位',
                    '非限制使用的库存': '数量',
                    '值未限制': '库存金额'
                }
                df_new = df_new.rename(columns=new_rename_dict)

                # 确保关键列是字符串格式，防止匹配失败 (去除浮点数末尾的 .0)
                df_new['物料编码'] = df_new['物料编码'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_old['物料编码'] = df_old['物料编码'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_new['批次'] = df_new['批次'].astype(str)
                df_old['批次'] = df_old['批次'].astype(str)

                # 5. 提取历史报表中的历史信息字典 (去除重复项以防报错)
                history_cols = ['物料编码', '批次', '采购订单', '入库时间', '供应商', '存放位置', '备注']
                # 如果旧表里有些字段本身就不存在，则忽略
                history_cols = [c for c in history_cols if c in df_old.columns]
                df_old_history = df_old[history_cols].drop_duplicates(subset=['物料编码', '批次'], keep='first')

                # 6. 核心逻辑：左连接 (Left Join)
                # 以新表为主表，去旧表里找对应的 订单、入库时间、位置 等信息
                df_merged = pd.merge(df_new, df_old_history, on=['物料编码', '批次'], how='left')

                # 7. 整理最终列的顺序
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

                # 9. 提供下载 (动态生成文件名)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='整合后库存明细')

                # 提取用户上传的新表名字（去掉后缀）
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
