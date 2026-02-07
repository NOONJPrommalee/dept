import win32com.client as win32
import os
import shutil

def convert_and_cleanup(source_folder):
    # 1. กำหนดโฟลเดอร์พักไฟล์ชั่วคราว
    archive_folder = os.path.join(source_folder, "Completed_Archive")
    if not os.path.exists(archive_folder):
        os.makedirs(archive_folder)

    # เปิด Excel Engine
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # 2. เริ่มกระบวนการแปลงไฟล์
        for filename in os.listdir(source_folder):
            if filename.lower().endswith(".xls") and not filename.startswith("~$"):
                xls_path = os.path.abspath(os.path.join(source_folder, filename))
                xlsx_path = xls_path + "x"

                print(f"กำลังจัดการ: {filename}...")
                
                # สั่งเปิดและ Save As
                wb = excel.Workbooks.Open(xls_path)
                wb.SaveAs(xlsx_path, FileFormat=51)
                wb.Close()

                # ย้ายไฟล์ที่ทำเสร็จแล้วไปไว้ใน Archive ก่อน
                shutil.move(xls_path, os.path.join(archive_folder, filename))
                print(f"แปลงไฟล์สำเร็จ: {filename}")

        # 🚩 3. ขั้นตอนการลบทิ้ง (Cleanup)
        # หลังจากวนลูปทำทุกไฟล์เสร็จแล้ว เราจะลบโฟลเดอร์ Archive ทิ้งทั้งหมด
        if os.path.exists(archive_folder):
            shutil.rmtree(archive_folder) # ลบโฟลเดอร์และไฟล์ข้างในทั้งหมด
            print("🧹 ล้างไฟล์ต้นฉบับใน Archive เรียบร้อยแล้ว")

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาด: {e}")
    finally:
        excel.Quit()
        print("✅ ปิดโปรแกรม Excel เรียบร้อย")

# ระบุโฟลเดอร์ที่เก็บไฟล์ .xls ของคุณ
convert_and_cleanup(r"D:\work\บน\dept\project_folder\convert")