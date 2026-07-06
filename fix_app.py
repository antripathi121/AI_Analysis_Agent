from pathlib import Path
import re

app_path = Path("app.py")
text = app_path.read_text(encoding="utf-8")

old = '''    row = 2
    while ws.cell(row, 1).value not in [None, ""]:
        row += 1
'''

new = '''    def normalize_key_value(value):
        if value is None:
            return ""
        try:
            return pd.to_datetime(value).strftime("%Y-%m-%d")
        except Exception:
            return str(value).strip().lower()

    target_row = None
    row = 2

    while ws.cell(row, 1).value not in [None, ""]:
        existing_cpg = str(ws.cell(row, 1).value).strip().lower()
        existing_offer = str(ws.cell(row, 2).value).strip().lower()
        existing_start = normalize_key_value(ws.cell(row, 5).value)
        existing_end = normalize_key_value(ws.cell(row, 6).value)

        if (
            existing_cpg == cpg_name_value.lower()
            and existing_offer == offer_name_value.lower()
            and existing_start == offer_start_value.strftime("%Y-%m-%d")
            and existing_end == offer_end_value.strftime("%Y-%m-%d")
        ):
            target_row = row
            break

        row += 1

    if target_row is None:
        target_row = row

    row = target_row
'''

count = text.count(old)

if count == 0:
    raise Exception("Could not find batch append row block.")

# Replace only the FIRST occurrence inside batch function if possible.
# This exact block should exist in save_batch_result_to_master.
text = text.replace(old, new, 1)

app_path.write_text(text, encoding="utf-8")

print("✅ Batch duplicate prevention added.")
print("✅ Existing CPG + Offer + Start + End will now update instead of duplicate.")