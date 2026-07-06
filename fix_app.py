from pathlib import Path

app_path = Path("app.py")
text = app_path.read_text(encoding="utf-8")

old_1 = 'sales_sheet = read_sales_sheet(sales_file)'
old_2 = 'sales_sheet = read_sales_sheet(str(sales_file))'

new = '''import tempfile
                import shutil

                temp_sales_file = Path(tempfile.gettempdir()) / f"batch_sales_{idx}.xlsx"
                shutil.copy2(str(sales_file), str(temp_sales_file))

                sales_sheet = read_sales_sheet(str(temp_sales_file))'''

if old_1 in text:
    text = text.replace(old_1, new)
elif old_2 in text:
    text = text.replace(old_2, new)
else:
    raise Exception("Could not find sales_sheet read line.")

app_path.write_text(text, encoding="utf-8")

print("✅ Batch file reading fixed using temporary local copy.")