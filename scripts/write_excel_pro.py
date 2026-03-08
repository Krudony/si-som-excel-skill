import xlwings as xw
import os
import sys

def write_pro(file_path, sheet_name, cell_ref, value, font_name='TH SarabunPSK', font_size=16):
    try:
        app = xw.App(visible=True) # Visible เพื่อความชัวร์ 100%
        app.display_alerts = False
        wb = app.books.open(os.path.abspath(file_path))
        ws = wb.sheets[sheet_name]
        
        target = ws.range(cell_ref)
        target.value = value
        target.api.Font.Name = font_name
        target.api.Font.Size = font_size
        target.api.HorizontalAlignment = -4108 # xlCenter
        
        wb.save()
        wb.close()
        app.quit()
        print(f"✅ Successfully wrote '{value}' to [{sheet_name}]{cell_ref} using xlwings.")
    except Exception as e:
        print(f"❌ xlwings Error: {str(e)}")
        try: app.quit()
        except: pass

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: python write_excel_pro.py <file> <sheet> <ref> <val>")
        sys.exit(1)
    write_pro(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
