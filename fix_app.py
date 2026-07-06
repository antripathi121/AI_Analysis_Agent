from pathlib import Path

app_path = Path("app.py")
text = app_path.read_text(encoding="utf-8")

old = '''run_batch = st.button("🚀 Run Batch Analysis")

from pathlib import Path
'''

new = '''batch_working_folder = st.text_input(
    "Local Batch Working Folder",
    value=r"C:\\Users\\ADMIN\\Desktop\\Batch Working"
)

prepare_batch = st.button("📂 Prepare Batch Folder")

run_batch = st.button("🚀 Run Batch Analysis")

from pathlib import Path
import shutil

if prepare_batch:
    try:
        source_folder = Path(batch_folder)
        target_folder = Path(batch_working_folder)

        if not source_folder.exists():
            st.error("Source Sales Report folder not found.")
            st.stop()

        target_folder.mkdir(parents=True, exist_ok=True)

        library_df = load_campaign_library()

        sales_files = [
            file for file in source_folder.glob("*.xlsx")
            if not file.name.startswith("~$")
        ]

        copied = []
        missing = []
        errors = []

        for _, campaign in library_df.iterrows():
            cpg_name = str(campaign.get("CPG", "")).strip()
            target_cpg = normalize_cpg_name(cpg_name)

            matched_file = None

            for file in sales_files:
                file_name_clean = normalize_cpg_name(file.stem)

                if file_name_clean == target_cpg:
                    matched_file = file
                    break

            if matched_file is None:
                for file in sales_files:
                    file_name_clean = normalize_cpg_name(file.stem)

                    if target_cpg in file_name_clean or file_name_clean in target_cpg:
                        matched_file = file
                        break

            if matched_file is None:
                missing.append({
                    "CPG": cpg_name,
                    "Reason": "No matching Sales Report filename found"
                })
                continue

            try:
                destination = target_folder / matched_file.name
                shutil.copy2(str(matched_file), str(destination))

                copied.append({
                    "CPG": cpg_name,
                    "Copied File": matched_file.name
                })

            except Exception as e:
                errors.append({
                    "CPG": cpg_name,
                    "Matched File": matched_file.name,
                    "Error": str(e)
                })

        st.success(f"Prepare complete. Copied: {len(copied)}, Missing: {len(missing)}, Errors: {len(errors)}")
        st.info(f"Prepared folder: {target_folder}")

        if copied:
            st.subheader("✅ Copied")
            st.dataframe(pd.DataFrame(copied), use_container_width=True)

        if missing:
            st.subheader("⚠️ Missing")
            st.dataframe(pd.DataFrame(missing), use_container_width=True)

        if errors:
            st.subheader("❌ Copy Errors")
            st.dataframe(pd.DataFrame(errors), use_container_width=True)

    except Exception as e:
        st.error(f"Prepare Batch Folder Error: {e}")

'''

if old not in text:
    raise Exception("Could not find Batch button block to replace.")

text = text.replace(old, new)

app_path.write_text(text, encoding="utf-8")

print("✅ Prepare Batch Folder button added.")