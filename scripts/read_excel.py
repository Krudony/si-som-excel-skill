import sys
import pandas as pd
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def analyze_excel(file_path):
    print(f"📊 Analyzing Excel File: {file_path}")
    print("-" * 50)
    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        print(f"📑 Total Sheets: {len(sheet_names)}")
        print(f"📝 Sheet Names: {', '.join(sheet_names)}\n")
        
        for sheet in sheet_names:
            print(f"--- Sheet: [{sheet}] ---")
            try:
                df = xl.parse(sheet)
                if df.empty:
                    print("⚠️ Sheet is empty.\n")
                    continue
                print(f"📐 Dimensions: {df.shape[0]} Rows x {df.shape[1]} Columns")
                print(f"🏷️ Headers: {', '.join(str(col) for col in df.columns.tolist())}")
                print("🔍 Data Sample (Top 3 rows):")
                print(df.head(3).to_markdown(index=False))
                print("\n")
            except Exception as e:
                print(f"❌ Error reading sheet '{sheet}': {str(e)}\n")
    except Exception as e:
        print(f"❌ Failed to open Excel file: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python read_excel.py <path_to_excel_file>")
        sys.exit(1)
    analyze_excel(sys.argv[1])
