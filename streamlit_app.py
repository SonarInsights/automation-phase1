import streamlit as st
import pandas as pd
import time
from datetime import datetime
import requests
from io import BytesIO

# === FUNGSI: Apply Rules ===
def apply_rules(df, rules, output_column, source_output_column):
    import re
    rules.columns = rules.columns.str.strip()
    rules_sorted = rules.sort_values(by="Priority", ascending=False)

    if output_column not in df.columns:
        df[output_column] = ""

    # Tambahkan: deteksi semua kolom output dari rules
    output_cols_in_rules = [col for col in rules.columns if col.startswith("Output ")]
    for output_col in output_cols_in_rules:
        colname = output_col.replace("Output ", "")
        if colname not in df.columns:
            df[colname] = ""
    priority_tracker = {
        col.replace("Output ", ""): pd.Series([float("inf")] * len(df), index=df.index)
        for col in output_cols_in_rules
    }

    summary_logs = []
    overwrite_tracker = [[] for _ in range(len(df))]

    for _, rule in rules_sorted.iterrows():
        col = rule["Matching Column"]
        val = rule["Matching Value"]
        match_type = rule["Matching Type"]
        priority = rule["Priority"]
        channel = rule.get("Channel", "")

        # Filter Channel
        if pd.notna(channel) and "Channel" in df.columns:
            channel_mask = df["Channel"].astype(str).str.lower() == str(channel).strip().lower()
        else:
            channel_mask = pd.Series([True] * len(df), index=df.index)

        # Matching logic (same as before)
        if "+" in col:
            parts = [p.strip() for p in col.split("+")]
            if not all(p in df.columns for p in parts):
                continue
            series = df[parts[0]].astype(str)
            for p in parts[1:]:
                series += "+" + df[p].astype(str)
        else:
            if col not in df.columns:
                continue
            series = df[col].astype(str)

        if match_type == "contains":
            if "+" in val:
                val_parts = val.split("+")
                submasks = []
                for v in val_parts:
                    v = v.strip()
                    if "|" in v:
                        keywords = [re.escape(x.strip()) for x in v.split("|")]
                        if all(k.startswith("\\!") for k in keywords):
                            keywords_clean = [k[2:] for k in keywords]
                            submask = ~series.str.contains("|".join(keywords_clean), case=False, na=False)
                        elif all(not k.startswith("\\!") for k in keywords):
                            submask = series.str.contains("|".join(keywords), case=False, na=False)
                        else:
                            must_not = [k[2:] for k in keywords if k.startswith("\\!")]
                            must_yes = [k for k in keywords if not k.startswith("\\!")]
                            mask_not = ~series.str.contains("|".join(must_not), case=False, na=False) if must_not else True
                            mask_yes = series.str.contains("|".join(must_yes), case=False, na=False) if must_yes else True
                            submask = mask_not & mask_yes
                    elif v.startswith("!"):
                        submask = ~series.str.contains(re.escape(v[1:]), case=False, na=False)
                    else:
                        submask = series.str.contains(re.escape(v), case=False, na=False)
                    submasks.append(submask)
                mask = pd.concat(submasks, axis=1).all(axis=1)
            else:
                if val.startswith("!"):
                    mask = ~series.str.contains(re.escape(val[1:]), case=False, na=False)
                else:
                    mask = series.str.contains(re.escape(val), case=False, na=False)
        elif match_type == "equals":
            mask = series == val
        elif match_type == "greater_than":
            try:
                val_num = float(val)
                series_num = pd.to_numeric(series, errors="coerce")
                mask = series_num > val_num
            except ValueError:
                continue
        elif match_type == "less_than":
            try:
                val_num = float(val)
                series_num = pd.to_numeric(series, errors="coerce")
                mask = series_num < val_num
            except ValueError:
                continue
        elif match_type == "count_contains":
            try:
                keyword, constraint = val.split(":")
                keyword = re.escape(keyword.strip())
                constraint = constraint.strip()
                counts = series.str.lower().str.count(rf"\b{keyword}\b")
                if "max=" in constraint:
                    max_allowed = int(constraint.replace("max=", "").strip())
                    mask = counts <= max_allowed
                elif "min=" in constraint:
                    min_allowed = int(constraint.replace("min=", "").strip())
                    mask = counts >= min_allowed
                else:
                    continue
            except Exception as e:
                print(f"⚠️ Error parsing count_contains rule: {val} - {e}")
                continue
        else:
            continue

        update_mask = mask & channel_mask

        # Apply output to all relevant output columns
        for output_col in output_cols_in_rules:
            out_val = rule.get(output_col)
            colname = output_col.replace("Output ", "")
            if pd.notna(out_val) and colname in df.columns:
                update_condition = update_mask & (priority_tracker[colname] > priority)
                df.loc[update_condition, colname] = out_val
                priority_tracker[colname].loc[update_condition] = priority
                for idx in update_condition[update_condition].index:
                    overwrite_tracker[idx].append(f"{colname} P{priority}: {out_val}")

        # Log
        if update_mask.sum() > 0:
            summary_logs.append({
                "Priority": priority,
                "Matching Column": col,
                "Matching Value": val,
                "Matching Type": match_type,
                "Channel": channel,
                "Affected Rows": update_mask.sum()
            })

    # Simpan chain overwrite jika hanya untuk Noise Tag
    df[output_column + " - Chain Overwrite"] = [" ➔ ".join(x) if x else "" for x in overwrite_tracker]
    summary_df = pd.DataFrame(summary_logs)
    return df, summary_df



#Untuk menentukan official account
def apply_official_account_logic(df, setup_df, project_name):
    import re
    setup_df.columns = setup_df.columns.str.strip()
    
    # Ubah nilai TRUE/FALSE jadi Yes/No (string)
    setup_df["Verified Account"] = setup_df["Verified Account"].apply(
        lambda x: "yes" if str(x).strip().lower() in ["true", "yes", "1"] else "no"
    )

    # Ambil rules yang sesuai project
    setup_project = setup_df[setup_df["Project"] == project_name]

    for _, row in setup_project.iterrows():
        verified = str(row.get("Verified Account", "")).strip().lower()
        channel = str(row.get("Channel", "")).strip().lower()
        col = row["Matching Column"]
        val = row["Matching Value"]
        match_type = row["Matching Type"]

        if col not in df.columns or "Channel" not in df.columns or "Verified Account" not in df.columns:
            continue

        # Filter: channel dan verified
        mask = (
            df["Verified Account"].astype(str).str.strip().str.lower() == verified
        ) & (
            df["Channel"].astype(str).str.strip().str.lower() == channel
        )

        series = df[col].astype(str)

        if match_type == "contains":
            pattern = re.escape(val)
            mask &= series.str.contains(pattern, case=False, na=False)

        elif match_type == "equals":
            mask &= series == val

        else:
            continue

        df.loc[mask, "Official Account"] = "Official Account"
        df.loc[mask, "Noise Tag"] = "1"

    return df



# === MULAI STREAMLIT APP ===
st.title("Insight Automation Phase 1")

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
        df_notes = pd.read_excel(xls, sheet_name="NOTES", header=None)
        last_updated = df_notes.iloc[0, 1]
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
    st.markdown("#### Pilih Project Name:")
    st.caption(f"📄 Rules terakhir diperbarui pada: {last_updated}")

    project_name = st.selectbox("", ["Pilih Project"] + df_project_list["Project Name"].dropna().tolist())

    uploaded_raw = st.file_uploader("Upload Raw Data", type=["xlsx"], key="raw")

    remove_duplicate_links = st.checkbox("Remove duplicate link")
    keep_raw_data = st.checkbox("Keep RAW Data (Save original file as separate sheet)")

    submit = st.button("Submit")

    if submit:
        if project_name == "Pilih Project" or uploaded_raw is None:
            st.error("❌ Anda harus memilih project dan upload raw data sebelum submit.")
        else:
            st.success(f"✅ Project: {project_name} | File Loaded Successfully!")

            start_time = time.time()

            df_raw = pd.read_excel(uploaded_raw, sheet_name=0)
            if "Campaign" in df_raw.columns:
                df_raw = df_raw.rename(columns={"Campaign": "Campaigns"})
            df_processed = df_raw.copy()

            # Remove duplicate link
            if remove_duplicate_links and "Link URL" in df_processed.columns:
                before_count = len(df_processed)
                df_processed = df_processed.drop_duplicates(subset="Link URL").reset_index(drop=True)
                after_count = len(df_processed)
                st.info(f"🔁 Removed {before_count - after_count} duplicate rows based on 'Link URL'")

            # Standardize Verified Account
            if "Verified Account" in df_processed.columns:
                df_processed["Verified Account"] = (
                    df_processed["Verified Account"].astype(str).str.strip().str.lower().replace({"-": "no", "": "no", "nan": "no"})
                )
                df_processed["Verified Account"] = df_processed["Verified Account"].apply(lambda x: "Yes" if x == "yes" else "No")


            # Setup Columns
            column_setup_default = df_column_setup[df_column_setup["Project"] == "Default"]
            column_setup_project = df_column_setup[df_column_setup["Project"] == project_name]
            column_setup_combined = pd.concat([column_setup_default, column_setup_project], ignore_index=True)

            for _, row in column_setup_combined.iterrows():
                col = row["Target Column"]
                ref_col = row["Reference Column"]
                pos = row["Position"]
                default = row["Default Value"] if pd.notna(row["Default Value"]) else ""

                if col not in df_processed.columns:
                    # Jika kolom belum ada, tambahkan dan isi default
                    if ref_col in df_processed.columns:
                        ref_idx = df_processed.columns.get_loc(ref_col)
                        insert_at = ref_idx if pos == "before" else ref_idx + 1
                        df_processed.insert(loc=insert_at, column=col, value=default)
                    else:
                        df_processed[col] = default
                else:
                    # Jika kolom sudah ada, isi semua nilai kosong / NaN dengan default
                    df_processed[col] = df_processed[col].fillna("").replace("", default)

            # Apply Official Account Logic dari setup sheet
            df_official_account_setup = pd.read_excel(xls, sheet_name="Official Account Setup")
            df_processed = apply_official_account_logic(df_processed, df_official_account_setup, project_name)


            # Bersihkan trailing .0 hanya untuk kolom 'Noise Tag' jika diperlukan
            if "Noise Tag" in df_processed.columns and df_processed["Noise Tag"].notna().any():
                try:
                    series_str = df_processed["Noise Tag"].astype(str)
                    if any(series_str.dropna().str.contains(r"\.0$", regex=True, na=False)):
                        df_processed["Noise Tag"] = series_str.replace({r"\.0$": ""}, regex=True)
                except Exception as e:
                    st.warning(f"⚠️ Gagal membersihkan kolom 'Noise Tag': {e}")

            # === Apply Rules ===
            rules_default = df_rules[df_rules["Project"] == "Default"]
            rules_project = df_rules[df_rules["Project"] == project_name] if project_name in df_rules["Project"].values else pd.DataFrame()
            rules_combined = pd.concat([rules_default, rules_project], ignore_index=True)

            # Apply untuk Noise Tag
            # Apply untuk Noise Tag
            df_processed, summary_df = apply_rules(
                df=df_processed,
                rules=rules_combined,
                output_column="Noise Tag",
                source_output_column="Output Noise Tag"
            )

            # Tambahkan ini untuk Issue
            df_processed, summary_df_issue = apply_rules(
                df=df_processed,
                rules=rules_combined,
                output_column="Issue",
                source_output_column="Output Issue"
            )

            df_processed, summary_df_sub_issue = apply_rules(
                df=df_processed,
                rules=rules_combined,
                output_column="Sub Issue",
                source_output_column="Output Sub Issue"
            )


            # Gabungkan summary Noise Tag + Issue + Sub Issue
            summary_combined = pd.concat([summary_df, summary_df_issue, summary_df_sub_issue], ignore_index=True)

            # === Hitung kolom Followers ===
            if "Original Reach" in df_processed.columns and "Potential Reach" in df_processed.columns:
                df_processed["Followers"] = df_processed["Original Reach"].fillna(0) + df_processed["Potential Reach"].fillna(0)


            # Setup Column Order
            if project_name in df_column_order["Project"].values:
                ordered_cols = df_column_order[df_column_order["Project"] == project_name]
            else:
                ordered_cols = df_column_order[df_column_order["Project"] == "Default"]

            ordered_cols = ordered_cols[ordered_cols["Hide"].str.lower() != "yes"]["Column Name"].tolist()
            final_cols = [col for col in ordered_cols if col in df_processed.columns]

            df_final = df_processed[final_cols]

            # Save Output
            tanggal_hari_ini = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"{project_name}_{tanggal_hari_ini}.xlsx"

            #Jika keep raw data dan tidak keep raw data
            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                if keep_raw_data:
                    df_raw.to_excel(writer, sheet_name="RAW Data", index=False)
                df_final.to_excel(writer, sheet_name="Process Data", index=False)


            end_time = time.time()
            minutes, seconds = divmod(end_time - start_time, 60)

            # === Hitung durasi proses
            duration_seconds = end_time - start_time
            hours, remainder = divmod(duration_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)

            st.info(f"🕒 Proses ini berjalan selama {int(hours)} jam {int(minutes)} menit {int(seconds)} detik.")


            # 1. Tampilkan Summary Execution Report
            st.subheader("📊 Summary Execution Report")
            with st.expander("Lihat Summary Execution Report"):
                if not summary_df.empty:
                     st.dataframe(summary_combined)
                else:
                    st.info("ℹ️ Tidak ada rule yang match pada data ini.")

            # 2. Tampilkan Chain Overwrite Tracker
            st.subheader("🧩 Chain Overwrite Tracker")
            with st.expander("Lihat Chain Overwrite Tracker"):
                chain_overwrite_columns = [output_column + " - Chain Overwrite" for output_column in ["Noise Tag"]]
                if any(col in df_final.columns for col in chain_overwrite_columns):
                    chain_overwrite_df = df_final[chain_overwrite_columns]
                    st.dataframe(chain_overwrite_df)
                else:
                    st.info("ℹ️ Tidak ada perubahan tercatat (Chain Overwrite kosong).")

            # 3. Tombol Download Hasil di paling bawah
            st.success(f"⏱️ Proses selesai dalam {int(minutes)} menit {int(seconds)} detik")
            st.download_button(
                label="⬇️ Download Hasil Excel",
                data=open(output_filename, "rb").read(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.stop()
