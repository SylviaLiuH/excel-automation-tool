
import io
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Excel 自动整理器",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Excel 自动整理器")
st.caption("上传两份 Excel / CSV 文件，选择关联字段，一键完成清洗、合并、缺失值检查和结果导出。")


def read_table(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)
    raise ValueError("目前只支持 .xlsx 和 .csv 文件")


def clean_dataframe(df):
    result = df.copy()

    # 清理列名
    result.columns = [str(c).strip() for c in result.columns]

    # 清理字符串字段首尾空格
    for col in result.select_dtypes(include="object").columns:
        result[col] = result[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )

    before = len(result)
    result = result.drop_duplicates()
    removed = before - len(result)

    return result, removed


def build_excel(merged_df, summary_df, missing_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="合并结果")
        summary_df.to_excel(writer, index=False, sheet_name="处理摘要")
        missing_df.to_excel(writer, index=False, sheet_name="缺失值报告")
    output.seek(0)
    return output.getvalue()


uploaded_1 = st.file_uploader("上传文件 A", type=["xlsx", "csv"], key="file_a")
uploaded_2 = st.file_uploader("上传文件 B", type=["xlsx", "csv"], key="file_b")

if uploaded_1 and uploaded_2:
    try:
        df1 = read_table(uploaded_1)
        df2 = read_table(uploaded_2)

        st.subheader("1️⃣ 数据预览")
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"**文件 A：{uploaded_1.name}**")
            st.dataframe(df1.head(8), use_container_width=True)
            st.caption(f"{len(df1)} 行 × {len(df1.columns)} 列")

        with col2:
            st.markdown(f"**文件 B：{uploaded_2.name}**")
            st.dataframe(df2.head(8), use_container_width=True)
            st.caption(f"{len(df2)} 行 × {len(df2.columns)} 列")

        common_columns = [c for c in df1.columns if c in df2.columns]

        if not common_columns:
            st.error("两份文件没有同名字段，暂时无法自动关联。请确保至少有一个共同字段，例如 order_id。")
            st.stop()

        st.subheader("2️⃣ 选择合并方式")

        join_col = st.selectbox(
            "选择关联字段",
            options=common_columns,
            help="例如两张表都有 order_id，就可以用 order_id 进行关联。",
        )

        join_mode_name = st.selectbox(
            "选择合并方式",
            options=["左连接（保留文件 A 全部数据）", "内连接（只保留两边都匹配的数据）", "外连接（保留两边全部数据）"],
        )

        join_mode_map = {
            "左连接（保留文件 A 全部数据）": "left",
            "内连接（只保留两边都匹配的数据）": "inner",
            "外连接（保留两边全部数据）": "outer",
        }

        if st.button("🚀 开始处理", type="primary", use_container_width=True):
            cleaned_1, removed_1 = clean_dataframe(df1)
            cleaned_2, removed_2 = clean_dataframe(df2)

            merged = pd.merge(
                cleaned_1,
                cleaned_2,
                how=join_mode_map[join_mode_name],
                on=join_col,
                suffixes=("_A", "_B"),
            )

            total_missing = int(merged.isna().sum().sum())
            missing_by_col = merged.isna().sum().reset_index()
            missing_by_col.columns = ["字段", "缺失值数量"]
            missing_by_col["缺失率"] = (
                missing_by_col["缺失值数量"] / max(len(merged), 1) * 100
            ).round(2).astype(str) + "%"

            summary = pd.DataFrame(
                {
                    "指标": [
                        "文件 A 原始行数",
                        "文件 B 原始行数",
                        "文件 A 删除重复行",
                        "文件 B 删除重复行",
                        "最终结果行数",
                        "最终结果列数",
                        "缺失值总数",
                        "关联字段",
                        "合并方式",
                        "处理时间",
                    ],
                    "结果": [
                        len(df1),
                        len(df2),
                        removed_1,
                        removed_2,
                        len(merged),
                        len(merged.columns),
                        total_missing,
                        join_col,
                        join_mode_name,
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ],
                }
            )

            st.success("处理完成！")

            a, b, c, d = st.columns(4)
            a.metric("结果行数", len(merged))
            b.metric("结果列数", len(merged.columns))
            c.metric("删除重复行", removed_1 + removed_2)
            d.metric("缺失值总数", total_missing)

            st.subheader("3️⃣ 合并结果")
            st.dataframe(merged, use_container_width=True)

            st.subheader("4️⃣ 缺失值报告")
            st.dataframe(missing_by_col, use_container_width=True)

            excel_bytes = build_excel(merged, summary, missing_by_col)

            st.download_button(
                "⬇️ 下载处理结果.xlsx",
                data=excel_bytes,
                file_name="processed_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    except Exception as e:
        st.exception(e)
else:
    st.info("先上传两份文件。可以直接用项目里的 sample_orders.xlsx 和 sample_customers.xlsx 测试。")
