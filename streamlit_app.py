import streamlit as st
import pandas as pd
import time
from datetime import datetime
import requests
from io import BytesIO

st.title("Insight Automation v5 - Streamlit Version")

# --- Load satu file Excel dari Google Drive (Project List + Rules) ---
try:
    file_id = "1qKZcRumDYft3SJ-Cl3qB65gwCRcB1rUZ"  # ID file Excel kamu

    # Download file Excel
    download_url = f"https://drive.google.com/uc?id={file_id}"
    response = requests.get(download_url)
    xls = pd.ExcelFile(BytesIO(response.content))

    # Load semua sheet yang dibutuhkan
    df_project_list = pd.read_excel(xls, sheet_name="Project List")
    df_column_setup = pd.read_excel(xls, sheet_name="Column Setup")
    df_rules = pd.read_excel(xls, sheet_name="Rules")
    df_column_order = pd.read_excel(xls, sheet_name="Column Order Setup")
    df_method_1_keyword = pd.read_excel(xls, sheet_name="Method 1 Keyword")
    df_method_selection = pd.read_excel(xls, sheet_name="Method Selection")

    # Load Last Updated dari NOTES!B2
    try:
        df_notes = pd.read_excel(xls, sheet_name="NOTES", header=None)  # tanpa header
        last_updated = df_notes.iloc[0, 1]  # baris ke-1, kolom ke-2 = B1
    except:
        last_updated = "Unknown"

    # Special Case Setup
    try:
        df_special_case = pd.read_excel(xls, sheet_name="Special Case Request")
    except:
        df_special_case = pd.DataFrame()

    load_success = True
except Exception as e:
    st.error(f"❌ Gagal load file dari Google Drive: {e}")
    load_success = False

if load_success:
    # Tampilkan informasi Project dan Last Update
    st.markdown("#### Pilih Project Name:")
    st.caption(f"📄 Rules terakhir diperbarui pada: {last_updated}")

    project_name = st.selectbox("", ["Pilih Project"] + df_project_list["Project Name"].dropna().tolist())

    uploaded_raw = st.file_uploader("Upload Raw Data", type=["xlsx"], key="raw")

    submit = st.button("Submit")

    if submit:
        if project_name == "Pilih Project" or uploaded_raw is None:
            st.error("❌ Anda harus memilih project dan upload raw data sebelum submit.")
        else:
            st.success(f"✅ Project: {project_name} | File Loaded Successfully!")

            start_time = time.time()

            # Load Raw Data
            #df_raw = pd.read_excel(uploaded_raw, sheet_name="Sheet1")
            df_raw = pd.read_excel(uploaded_raw, sheet_name=0) #baca sheet pertama
            
            # Cek dan ganti kolom 'Campaign' menjadi 'Campaigns'
            if "Campaign" in df_raw.columns:
                df_raw = df_raw.rename(columns={"Campaign": "Campaigns"})
            
            df_processed = df_raw.copy()

            # Standardize Verified Account
            if "Verified Account" in df_processed.columns:
                df_processed["Verified Account"] = (
                    df_processed["Verified Account"].astype(str).str.strip().str.lower().replace({"-": "no", "": "no", "nan": "no"})
                )
                df_processed["Verified Account"] = df_processed["Verified Account"].apply(lambda x: "Yes" if x == "yes" else "No")

            # Setup Column
            column_setup = df_column_setup[df_column_setup["Project"] == project_name]
            for _, row in column_setup.iterrows():
                col, ref_col, pos, default = row["Target Column"], row["Reference Column"], row["Position"], row["Default Value"]
                if col not in df_processed.columns:
                    if ref_col in df_processed.columns:
                        ref_idx = df_processed.columns.get_loc(ref_col)
                        insert_at = ref_idx if pos == "before" else ref_idx + 1
                        df_processed.insert(loc=insert_at, column=col, value=default)
                    else:
                        df_processed[col] = default

            # Init default columns Official Account and Noise Tag if not adding by rules
            '''for col in ["Official Account", "Noise Tag"]:
                if col not in df_processed.columns:
                    df_processed[col] = ""'''

            #Hilangkan jika ada .0 di data Noise Tag dan Official Account
            df_processed["Noise Tag"] = df_processed["Noise Tag"].replace({".0": ""}, regex=True)
            df_processed["Official Account"] = df_processed["Official Account"].replace({".0": ""}, regex=True)

            # Apply Rules
            rules_sorted = df_rules[df_rules["Project"] == project_name].sort_values(by="Priority", ascending=False)
            priority_tracker = {col: pd.Series([float("inf")] * len(df_processed), index=df_processed.index) for col in ["Noise Tag", "Official Account"]}

            for _, rule in rules_sorted.iterrows():
                col = rule["Matching Column"]
                val = rule["Matching Value"]
                match_type = rule["Matching Type"]
                priority_outputs = {}

                if str(rule.get("Affects Noise Tag", "")).strip().lower() == "yes":
                    priority_outputs["Noise Tag"] = rule["Output Noise Tag"]
                if str(rule.get("Affects Official Account", "")).strip().lower() == "yes":
                    priority_outputs["Official Account"] = rule["Output Official Account"]

                if col not in df_processed.columns:
                    continue

                series = df_processed[col].astype(str)
                if match_type == "contains":
                    mask = series.str.contains(val, case=False, na=False)
                elif match_type == "equals":
                    mask = series == val
                else:
                    continue

                for out_col, out_val in priority_outputs.items():
                    if out_col not in df_processed.columns:
                        df_processed[out_col] = ""
                    update_mask = mask & (priority_tracker[out_col] > rule["Priority"])
                    df_processed.loc[update_mask, out_col] = out_val
                    priority_tracker[out_col].loc[update_mask] = rule["Priority"]

            # Column Order
            ordered_cols = df_column_order[df_column_order["Project"] == project_name]
            ordered_cols = ordered_cols[ordered_cols["Hide"].str.lower() != "yes"]["Column Name"].tolist()
            final_cols = [col for col in ordered_cols if col in df_processed.columns]

            df_final = df_processed[final_cols]

            # Save Output
            tanggal_hari_ini = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"{project_name}_{tanggal_hari_ini}.xlsx"
            df_final.to_excel(output_filename, index=False)

            end_time = time.time()
            minutes, seconds = divmod(end_time - start_time, 60)

            st.success(f"⏱️ Proses selesai dalam {int(minutes)} menit {int(seconds)} detik")
            st.download_button(
                label="Download Hasil",
                data=open(output_filename, "rb").read(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.stop()
