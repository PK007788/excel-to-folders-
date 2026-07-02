import pandas as pd

excel_path = r"C:\Users\prajn\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions\AAE1574AAF895F5ABB04088C9F056BE2033BCF10\transfers\2026-27\updated_retino (1).xlsx"

print("Inspecting headers...")
xls = pd.ExcelFile(excel_path)
for sheet in xls.sheet_names:
    df = pd.read_excel(excel_path, sheet_name=sheet, nrows=5)
    print(f"Sheet: {sheet}")
    print("Columns:", list(df.columns))
    if "Retinopathy grade" in df.columns:
        print("Grades:", df["Retinopathy grade"].tolist())
    else:
        print("Grades column NOT found!")
    print("-" * 40)
