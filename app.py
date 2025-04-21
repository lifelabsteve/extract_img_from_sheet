import time
from pathlib import Path
from PIL import ImageGrab
from win32com.client import DispatchEx

def save_all_sdr_sheets_as_images():
    excel_filename = "Design Rules.xlsx"
    range_address = "B7:N39"

    current_dir = Path(__file__).parent.resolve()
    excel_path = current_dir / excel_filename
    output_dir = current_dir / "output"
    output_dir.mkdir(exist_ok=True)

    if not excel_path.exists():
        print(f"❌ 엑셀 파일이 존재하지 않습니다:\n{excel_path}")
        return

    print(f"✅ 엑셀 파일 경로: {excel_path}")

    excel = DispatchEx("Excel.Application")
    excel.Visible = False

    wb = None
    try:
        wb = excel.Workbooks.Open(str(excel_path))

        for sheet in wb.Sheets:
            sheet_name = sheet.Name
            if sheet_name.startswith("SDR"):
                output_path = output_dir / f"{sheet_name}.png"
                print(f"🖼 시트 처리 중: {sheet_name} → {output_path}")

                # ✅ 고해상도 이미지 위해 줌 확대
                sheet.Activate()
                excel.ActiveWindow.Zoom = 200

                sheet.Range(range_address).CopyPicture(Appearance=1, Format=2)
                time.sleep(1)

                image = ImageGrab.grabclipboard()
                if image is None:
                    print(f"❌ 이미지 복사 실패: {sheet_name}")
                else:
                    image.save(output_path)
                    print(f"✅ 저장 완료: {output_path}")

    except Exception as e:
        print("❌ Excel 처리 중 오류:", e)
    finally:
        if wb:
            wb.Close(False)
        excel.Quit()

if __name__ == "__main__":
    save_all_sdr_sheets_as_images()
