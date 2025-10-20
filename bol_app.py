import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io, zipfile, os

# ================= 基本設定 =================
TEMPLATE_PATH = "BOL.pdf"  # 固定從 repo 根目錄讀取
st.set_page_config(page_title="BOL 自動產生器", page_icon="📦", layout="wide")

# ================= 工具函數 =================
def sanitize_token(s: str, max_len: int | None = None) -> str:
    if not s:
        return ""
    s = str(s)
    if max_len:
        s = s[:max_len]
    s = s.replace(" ", "_").strip()
    return "".join(ch for ch in s if ch.isalnum() or ch in "._-")

def first_n_alnum(s: str, n: int) -> str:
    if not s:
        return ""
    return "".join(ch for ch in s if ch.isalnum())[:n]

def make_output_name(row: dict, index: int) -> str:
    bol = sanitize_token(str(row.get("BOLnum", "")).strip())
    if not bol:
        bol = f"ROW_{index+1:03d}"
    desc8 = sanitize_token(str(row.get("Desc_1", "")).strip(), max_len=8)
    from2 = first_n_alnum(str(row.get("FromName", "")).strip(), 2)
    scac = sanitize_token(str(row.get("SCAC", "")).strip())
    parts = ["BOL", bol]
    if desc8: parts.append(desc8)
    if from2: parts.append(from2)
    if scac: parts.append(scac)
    return "_".join(parts) + ".pdf"

def fill_pdf(template_pdf: bytes, row: dict, index: int) -> bytes:
    # 依你的 PDF 表單欄位命名調整
    row = dict(row)
    row["3rdParty"] = "X"
    row["PrePaid"] = ""
    row["Collect"] = ""

    doc = fitz.open("pdf", template_pdf)
    for page in doc:
        widgets = page.widgets() or []
        for w in widgets:
            fname = w.field_name
            if not fname or fname not in row:
                continue
            try:
                w.field_value = str(row[fname])
                w.update()
            except Exception:
                pass
    try:
        doc.need_appearances = True
    except Exception:
        pass
    pdf_bytes = doc.tobytes(deflate=True)
    doc.close()
    return pdf_bytes

@st.cache_data(show_spinner=False)
def load_repo_template(path: str) -> bytes:
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到內建模板：{path}")
    with open(path, "rb") as f:
        data = f.read()
    if not data.startswith(b"%PDF"):
        raise ValueError("內建模板不是有效 PDF")
    return data

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def to_json_bytes(df: pd.DataFrame) -> bytes:
    return df.to_json(orient="records", force_ascii=False, indent=2).encode("utf-8")

# ================= UI =================
st.title("📦 BOL 自動產生器（固定使用 repo 中的 BOL.pdf）")
st.caption("模板來源：repo main 根目錄的 `BOL.pdf`。BOL 會對 Excel 的 **全部列** 產生；下方勾選只影響 CSV/JSON 匯出（供其他 API 用）。")

# 載入固定模板
try:
    template_bytes = load_repo_template(TEMPLATE_PATH)
    st.success(f"已載入模板：`{TEMPLATE_PATH}`")
except Exception as e:
    st.error(f"模板載入失敗：{e}")
    st.stop()

# 上傳 Excel
excel_file = st.file_uploader("上傳 Excel（.xlsx）", type=["xlsx"], help="BOL 將針對全部列產生；你可勾選部分列匯出 CSV/JSON 供其他 API 使用。")

if excel_file:
    try:
        df_all = pd.read_excel(excel_file, dtype=object).fillna("")
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}")
        st.stop()

    st.write(f"已載入 {len(df_all)} 筆資料。")
    with st.expander("🔎 檢視前 5 筆資料（核對欄位）", expanded=False):
        st.dataframe(df_all.head(5), use_container_width=True)

    # ------ BOL 產生（永遠針對全部列） ------
    st.subheader("🧾 產生所有列的 BOL PDF（ZIP）")
    if st.button("開始產生 BOL（全部列）", type="primary"):
        with st.spinner("正在生成 PDF..."):
            tmp_zip = io.BytesIO()
            with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, row in df_all.iterrows():
                    outname = make_output_name(row, i)
                    filled_pdf = fill_pdf(template_bytes, row.to_dict(), i)
                    zf.writestr(outname, filled_pdf)
            tmp_zip.seek(0)
        st.success(f"完成！已輸出 {len(df_all)} 份 BOL。")
        st.download_button(
            "📥 下載所有 BOL（ZIP）",
            tmp_zip,
            file_name="BOL_All.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.divider()

    # ------ 勾選列（只用於 API 匯出） ------
    st.subheader("✅ 勾選要送 API 的列（不影響 BOL 產生）")

    # 初始化可編輯表（預設全選）
    if "api_rows" not in st.session_state or st.session_state.get("data_version_key") != excel_file.name:
        api_rows = df_all.copy()
        api_rows.insert(0, "use", True)
        st.session_state.api_rows = api_rows
        st.session_state.data_version_key = excel_file.name

    edited_df = st.data_editor(
        st.session_state.api_rows,
        key="api_rows_editor",
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config={
            "use": st.column_config.CheckboxColumn("使用", help="勾選表示該列會被匯出到 CSV/JSON 供 API 使用", default=True),
        },
    )
    st.session_state.api_rows = edited_df

    selected_df = edited_df[edited_df["use"] == True].drop(columns=["use"])
    st.info(f"目前勾選 {len(selected_df)} / {len(edited_df)} 列供 API 使用（與 BOL 產生無關）。")

    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button(
            "下載勾選列（CSV）",
            to_csv_bytes(selected_df),
            file_name="api_selected_rows.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col_b:
        st.download_button(
            "下載勾選列（JSON）",
            to_json_bytes(selected_df),
            file_name="api_selected_rows.json",
            mime="application/json",
            use_container_width=True,
        )

    # 欄位提示（不會擋）
    required_cols = ["BOLnum", "Desc_1", "FromName", "SCAC"]
    missing_all = [c for c in required_cols if c not in df_all.columns]
    if missing_all:
        st.warning(f"Excel 缺少常用欄位：{missing_all}（仍可生成 BOL，但檔名/欄位可能不完整）")
else:
    st.info("請先上傳 Excel。")
