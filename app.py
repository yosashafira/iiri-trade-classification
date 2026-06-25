import os
import zipfile
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st
from dbfread import DBF

try:
    import rarfile
except Exception:
    rarfile = None


st.set_page_config(
    page_title="IIRI Trade Classification",
    page_icon="📊",
    layout="wide"
)

st.title("📊 IIRI Trade Classification")
st.caption("Aplikasi pemetaan data ekspor-impor DBF ke ISIC dan intensitas teknologi")

st.divider()


def clean_hs(x):
    if pd.isna(x):
        return ""
    return str(x).replace(".0", "").strip().zfill(8)


def extract_uploaded_files(uploaded_files, temp_dir):
    dbf_paths = []

    for uploaded_file in uploaded_files:
        file_path = os.path.join(temp_dir, uploaded_file.name)

        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        lower_name = uploaded_file.name.lower()

        if lower_name.endswith(".dbf"):
            dbf_paths.append(file_path)

        elif lower_name.endswith(".zip"):
            extract_dir = os.path.join(temp_dir, uploaded_file.name + "_unzipped")
            os.makedirs(extract_dir, exist_ok=True)

            with zipfile.ZipFile(file_path, "r") as zip_ref:
                zip_ref.extractall(extract_dir)

            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.lower().endswith(".dbf"):
                        dbf_paths.append(os.path.join(root, file))

        elif lower_name.endswith(".rar"):
            if rarfile is None:
                st.warning("RAR belum bisa dibaca. Untuk trial besok, gunakan ZIP atau DBF langsung.")
                continue

            extract_dir = os.path.join(temp_dir, uploaded_file.name + "_unrar")
            os.makedirs(extract_dir, exist_ok=True)

            try:
                with rarfile.RarFile(file_path, "r") as rar_ref:
                    rar_ref.extractall(extract_dir)

                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.lower().endswith(".dbf"):
                            dbf_paths.append(os.path.join(root, file))

            except Exception as e:
                st.warning(f"RAR gagal dibaca: {uploaded_file.name}. Untuk trial, pakai ZIP dulu.")

    return dbf_paths


def process_dbf_to_agg(dbf_path):
    file_name = os.path.basename(dbf_path)

    table = DBF(dbf_path, load=False, encoding="latin1")

    chunk_size = 100000
    temp_chunk = []
    chunks = []

    for i, record in enumerate(table):
        temp_chunk.append(record)

        if (i + 1) % chunk_size == 0:
            df_chunk = pd.DataFrame(temp_chunk)
            df_chunk.columns = df_chunk.columns.str.lower()

            df_chunk["val"] = pd.to_numeric(df_chunk["val"], errors="coerce")
            df_chunk["netwt"] = pd.to_numeric(df_chunk["netwt"], errors="coerce")

            agg = df_chunk.groupby(["hscode", "descr", "sitc"]).agg(
                frekuensi=("hscode", "count"),
                total_val=("val", "sum"),
                total_netwt=("netwt", "sum")
            ).reset_index()

            chunks.append(agg)
            temp_chunk = []

    if temp_chunk:
        df_chunk = pd.DataFrame(temp_chunk)
        df_chunk.columns = df_chunk.columns.str.lower()

        df_chunk["val"] = pd.to_numeric(df_chunk["val"], errors="coerce")
        df_chunk["netwt"] = pd.to_numeric(df_chunk["netwt"], errors="coerce")

        agg = df_chunk.groupby(["hscode", "descr", "sitc"]).agg(
            frekuensi=("hscode", "count"),
            total_val=("val", "sum"),
            total_netwt=("netwt", "sum")
        ).reset_index()

        chunks.append(agg)

    df_final = pd.concat(chunks, ignore_index=True)

    df_final = df_final.groupby(["hscode", "descr", "sitc"]).agg(
        frekuensi=("frekuensi", "sum"),
        total_val=("total_val", "sum"),
        total_netwt=("total_netwt", "sum")
    ).reset_index()

    name_lower = file_name.lower()

    if "ekspor" in name_lower:
        trade_type = "ekspor"
    elif "impor" in name_lower:
        trade_type = "impor"
    else:
        trade_type = "unknown"

    year = "".join([c for c in file_name if c.isdigit()])
    year = year[-4:] if len(year) >= 4 else ""

    df_final["year"] = year
    df_final["trade_type"] = trade_type
    df_final["source_file"] = file_name

    return df_final


def prepare_corres(corres_file):
    corres = pd.read_excel(corres_file, dtype=str)
    corres.columns = corres.columns.str.strip()

    required_cols = [
        "HS_8",
        "Description HS_8",
        "SITC Final",
        "ISIC Rev4.0",
        "Description ISIC Rev4.0",
        "Manufaktur atau Non Manufaktur",
        "Intensitas IIRI-OECD (5 kategori)",
        "Intensitas IIRI-UN (4 Kategori)",
        "Intensitas BPS",
        "Reason 2026"
    ]

    missing = [col for col in required_cols if col not in corres.columns]

    if missing:
        raise ValueError(f"Kolom ini tidak ditemukan di Final Corres: {missing}")

    corres["HS_8_clean"] = corres["HS_8"].apply(clean_hs)

    if corres["HS_8_clean"].duplicated().any():
        raise ValueError("Ada HS_8 yang duplikat di Final Corres. Cek file korespondensi dulu.")

    corres_keep = corres[[
        "HS_8_clean",
        "HS_8",
        "Description HS_8",
        "SITC Final",
        "ISIC Rev4.0",
        "Description ISIC Rev4.0",
        "Manufaktur atau Non Manufaktur",
        "Intensitas IIRI-OECD (5 kategori)",
        "Intensitas IIRI-UN (4 Kategori)",
        "Intensitas BPS",
        "Reason 2026"
    ]]

    return corres_keep


def process_all(trade_files, corres_file):
    output = BytesIO()
    summary = []

    with tempfile.TemporaryDirectory() as temp_dir:
        dbf_paths = extract_uploaded_files(trade_files, temp_dir)

        if len(dbf_paths) == 0:
            raise ValueError("Tidak ada file DBF yang ditemukan.")

        corres_keep = prepare_corres(corres_file)

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            progress = st.progress(0)
            status = st.empty()

            total_files = len(dbf_paths)

            for idx, dbf_path in enumerate(dbf_paths, start=1):
                file_name = os.path.basename(dbf_path)

                status.write(f"Memproses {file_name}...")

                df = process_dbf_to_agg(dbf_path)
                df["hscode_clean"] = df["hscode"].apply(clean_hs)

                mapped = df.merge(
                    corres_keep,
                    left_on="hscode_clean",
                    right_on="HS_8_clean",
                    how="left"
                )

                mapped = mapped.drop(columns=["hscode_clean", "HS_8_clean"])

                total_baris = len(mapped)
                termapping = mapped["ISIC Rev4.0"].notna().sum()
                unmapped = mapped["ISIC Rev4.0"].isna().sum()
                persen = round(termapping / total_baris * 100, 2) if total_baris > 0 else 0

                summary.append({
                    "file": file_name,
                    "year": mapped["year"].iloc[0] if total_baris > 0 else "",
                    "trade_type": mapped["trade_type"].iloc[0] if total_baris > 0 else "",
                    "total_baris": total_baris,
                    "termapping": termapping,
                    "belum_termapping": unmapped,
                    "persen_termapping": persen,
                    "total_frekuensi": mapped["frekuensi"].sum(),
                    "total_val": mapped["total_val"].sum(),
                    "total_netwt": mapped["total_netwt"].sum()
                })

                sheet_name = file_name.replace(".dbf", "").replace(".DBF", "")[:31]
                mapped.to_excel(writer, sheet_name=sheet_name, index=False)

                progress.progress(idx / total_files)

            summary_df = pd.DataFrame(summary)
            summary_df.to_excel(writer, sheet_name="summary_mapping", index=False)

            status.write("Selesai.")

    output.seek(0)
    return output, pd.DataFrame(summary)


st.header("① Upload Data Perdagangan")
trade_files = st.file_uploader(
    "Upload file DBF, ZIP, atau RAR",
    type=["dbf", "zip", "rar"],
    accept_multiple_files=True
)

st.header("② Upload File Korespondensi")
corres_file = st.file_uploader(
    "Upload Final Corres.xlsx",
    type=["xlsx"]
)

st.header("③ Proses")
process = st.button("🚀 Mulai Proses", use_container_width=True)

if process:
    if not trade_files:
        st.error("Upload file DBF/ZIP/RAR dulu.")
    elif corres_file is None:
        st.error("Upload Final Corres.xlsx dulu.")
    else:
        try:
            with st.spinner("Sedang memproses data. Mohon tunggu..."):
                result_file, summary_df = process_all(trade_files, corres_file)

            st.success("Proses selesai.")

            st.subheader("Ringkasan Mapping")
            st.dataframe(summary_df, use_container_width=True)

            st.download_button(
                label="📥 Download Hasil Excel",
                data=result_file,
                file_name="Data Ekspor Impor Mapped.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error("Proses gagal.")
            st.exception(e)