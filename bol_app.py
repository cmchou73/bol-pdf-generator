import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io, os, zipfile, tempfile, re

# --------------- 工具函數 ---------------
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


# --------------- Streamlit 網頁介面 ---------------
st.title("📦 BOL 自動產生器")
st.markdown("上傳 Excel（含 BOLnum、Desc_1、FromName、SCAC 等欄位）和 BOL 模板 PDF，批次產生帶欄位的 BOL PDF。")

template_file = st.file_uploader("上傳 BOL 模板 (PDF)", type=["pdf"])
excel_file = st.file_uploader("上傳 Excel (xlsx)", type=["xlsx"])

if template_file and excel_file:
    df = pd.read_excel(excel_file, dtype=object).fillna("")
    st.write(f"共 {len(df)} 筆資料。")

    if st.button("開始產生 BOL PDF"):
        with st.spinner("正在生成 PDF..."):
            pdf_bytes = template_file.read()
            tmp_zip = io.BytesIO()
            with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, row in df.iterrows():
                    outname = make_output_name(row, i)
                    filled_pdf = fill_pdf(pdf_bytes, row.to_dict(), i)
                    zf.writestr(outname, filled_pdf)
            tmp_zip.seek(0)
        st.success("✅ 完成！")
        st.download_button("📥 下載所有 BOL（ZIP）", tmp_zip, file_name="BOL_All.zip", mime="application/zip")
