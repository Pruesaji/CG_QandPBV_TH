# ============================================================
# สรุปข้อมูล Meeting Attendance จากไฟล์ CSV ทุกตัว
# ============================================================
# วิธีรัน:
#   python aggregate_attendance.py
#
# Input:
#   โฟลเดอร์ output/ ที่มีไฟล์ CSV จาก setsmart_login.py เช่น:
#     output/2S_meeting_attendance.csv
#     output/A_meeting_attendance.csv
#
# Output:
#   output/summary_board_attendance.csv
# ============================================================

import csv
import os
import glob

# ============================================================
# ตั้งค่า
# ============================================================
INPUT_DIR = "output"                                    # โฟลเดอร์ที่เก็บไฟล์ CSV รายบริษัท
OUTPUT_FILE = os.path.join(INPUT_DIR, "summary_board_attendance.csv")  # ไฟล์สรุป

# ============================================================
# ฟังก์ชันหลัก
# ============================================================

def parse_meeting_fraction(text):
    """
    แปลงข้อความ "24 / 24" หรือ "4/4" ให้เป็นตัวเลข 2 ตัว: (attended, total)

    ตัวอย่าง:
      "24 / 24"  → (24, 24)
      "4 / 5"    → (4, 5)
      "0 / 0"    → (0, 0)
      "N/A"      → None  (ไม่สามารถแปลงได้)
      ""         → None

    วิธีทำงาน:
      1. ลบช่องว่างออก → "24/24"
      2. แยกด้วย "/" → ["24", "24"]
      3. แปลงเป็นตัวเลข → (24, 24)
    """
    if not text or text.strip() == "" or text.strip() == "N/A":
        return None

    # ลบช่องว่าง
    cleaned = text.replace(" ", "")

    # แยกด้วย /
    parts = cleaned.split("/")

    if len(parts) != 2:
        return None

    try:
        attended = int(parts[0])   # ตัวเลขก่อน / = จำนวนครั้งที่เข้าประชุมจริง
        total = int(parts[1])      # ตัวเลขหลัง / = จำนวนครั้งที่ควรเข้าประชุม
        return (attended, total)
    except ValueError:
        return None


def process_one_file(filepath):
    """
    อ่านไฟล์ CSV 1 ไฟล์ แล้วคำนวณสรุปแยกตามปี

    ขั้นตอน:
      1. อ่านทุกแถว
      2. กรองเฉพาะ Section = "Meeting Attendance of the Board of Directors"
         (ไม่เอา Audit Committee, Executive Committee ฯลฯ)
      3. แยกข้อมูลตามปี
      4. คำนวณ:
         - Total_Board_Meetings_Per_Year
         - Company_Overall_Attendance
         - Company_Overall_Attendance_Percentage

    Returns:
      list ของ dict เช่น:
      [
        {"Company_Symbol": "2S", "Year": 2021,
         "Total_Board_Meetings_Per_Year": 5,
         "Company_Overall_Attendance": 49,
         "Company_Overall_Attendance_Percentage": 98.00},
        ...
      ]
    """
    results = []

    # ----- อ่านไฟล์ CSV -----
    # encoding="utf-8-sig" เพราะไฟล์ถูกบันทึกด้วย BOM
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    if not rows:
        return results

    # ดึง symbol จากแถวแรก
    symbol = rows[0].get("Symbol", "")
    if not symbol:
        # ลองดึงจากชื่อไฟล์ เช่น "2S_meeting_attendance.csv" → "2S"
        basename = os.path.basename(filepath)
        symbol = basename.replace("_meeting_attendance.csv", "")

    # ----- กรองเฉพาะ Board of Directors -----
    board_rows_all = [
        r for r in rows
        if "Board of Directors" in r.get("Section", "")
    ]

    # แยก: แถวที่มีข้อมูลจริง vs แถวที่เป็น No Information
    board_rows = [
        r for r in board_rows_all
        if r.get("Name", "") not in ("No Information", "No Data", "")
    ]

    if not board_rows_all:
        return results

    # ----- รวบรวมทุกปีที่มีใน CSV (รวม No Information) -----
    all_years = set()
    for r in board_rows_all:
        try:
            all_years.add(int(r["Year"]))
        except (ValueError, KeyError):
            pass

    # ----- จัดกลุ่มแถวที่มีข้อมูลจริง ตามปี -----
    year_data = {}
    for r in board_rows:
        try:
            year = int(r["Year"])
        except (ValueError, KeyError):
            continue
        if year not in year_data:
            year_data[year] = []
        year_data[year].append(r)

    # ----- คำนวณแต่ละปี -----
    for year in sorted(all_years):

        # ---- ปีที่ไม่มีข้อมูลจริง → ใส่ N/A ----
        if year not in year_data:
            results.append({
                "Company_Symbol": symbol,
                "Year": year,
                "Total_Board_Meetings_Per_Year": "N/A",
                "Company_Overall_Attendance": "N/A",
                "Total_Should_Attend": "N/A",
                "Company_Overall_Attendance_Percentage": "N/A",
            })
            continue

        members = year_data[year]

        total_attended = 0
        total_should_attend = 0
        board_meetings_per_year = 0
        valid_members = 0

        # ---- ลองใช้ Board Meeting Count จาก CSV ก่อน ----
        csv_board_count = members[0].get("Board Meeting Count", "").strip()
        if csv_board_count and csv_board_count != "N/A":
            try:
                board_meetings_per_year = int(csv_board_count)
            except ValueError:
                pass

        for member in members:
            fraction = parse_meeting_fraction(member.get("Number of Board Meeting", ""))

            if fraction is None:
                continue

            attended, should_attend = fraction

            # ---- ข้าม row ที่ 0/0 ----
            if attended == 0 and should_attend == 0:
                continue

            total_attended += attended
            total_should_attend += should_attend
            valid_members += 1

            # ---- หา Total_Board_Meetings_Per_Year (fallback) ----
            if board_meetings_per_year == 0:
                term_date = member.get("Termination Date", "").strip()
                if term_date in ("N/A", "", "-"):
                    board_meetings_per_year = max(board_meetings_per_year, should_attend)

        # ---- ถ้ามีแถวข้อมูลแต่ทุกคนเป็น 0/0 → ใส่ N/A ----
        if valid_members == 0:
            results.append({
                "Company_Symbol": symbol,
                "Year": year,
                "Total_Board_Meetings_Per_Year": "N/A",
                "Company_Overall_Attendance": "N/A",
                "Total_Should_Attend": "N/A",
                "Company_Overall_Attendance_Percentage": "N/A",
            })
            continue

        # ----- คำนวณ Percentage -----
        if total_should_attend > 0:
            percentage = round((total_attended / total_should_attend) * 100, 2)
        else:
            percentage = 0.0

        results.append({
            "Company_Symbol": symbol,
            "Year": year,
            "Total_Board_Meetings_Per_Year": board_meetings_per_year,
            "Company_Overall_Attendance": total_attended,
            "Total_Should_Attend": total_should_attend,
            "Company_Overall_Attendance_Percentage": percentage,
        })

    return results


def main():
    """
    ฟังก์ชันหลัก:
      1. หาไฟล์ CSV ทุกไฟล์ในโฟลเดอร์ output/
      2. ประมวลผลทีละไฟล์
      3. รวมผลลัพธ์ทั้งหมดแล้วบันทึกเป็น CSV ไฟล์เดียว
    """

    # ----- หาไฟล์ CSV ทั้งหมด -----
    # glob.glob() = หาไฟล์ที่ตรงกับ pattern
    # "*_meeting_attendance.csv" = ชื่ออะไรก็ได้ ที่ลงท้ายด้วย _meeting_attendance.csv
    pattern = os.path.join(INPUT_DIR, "*_meeting_attendance.csv")
    csv_files = sorted(glob.glob(pattern))

    if not csv_files:
        print(f"❌ ไม่เจอไฟล์ CSV ในโฟลเดอร์ {INPUT_DIR}/")
        print(f"   กรุณารัน setsmart_login.py ก่อน")
        return

    print(f"📁 เจอไฟล์ CSV {len(csv_files)} ไฟล์:")
    for f in csv_files:
        print(f"   - {os.path.basename(f)}")

    # ----- ประมวลผลทีละไฟล์ -----
    all_results = []

    for filepath in csv_files:
        basename = os.path.basename(filepath)
        print(f"\n📊 กำลังประมวลผล: {basename}")

        file_results = process_one_file(filepath)

        if file_results:
            all_results.extend(file_results)
            for r in file_results:
                print(f"   year {r['Year']}: "
                      f"ประชุม {r['Total_Board_Meetings_Per_Year']} ครั้ง, "
                      f"เข้าร่วมรวม {r['Company_Overall_Attendance']}/{r['Total_Should_Attend']} ครั้ง, "
                      f"= {r['Company_Overall_Attendance_Percentage']}%")
        else:
            print(f"   ⚠️ No data available for Board of Directors")

    # ----- บันทึกไฟล์สรุป -----
    if not all_results:
        print("\n❌ No data to save!")
        return

    fieldnames = [
        "Company_Symbol",
        "Year",
        "Total_Board_Meetings_Per_Year",
        "Company_Overall_Attendance",
        "Total_Should_Attend",
        "Company_Overall_Attendance_Percentage",
    ]

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_results)

    print(f"\n{'='*50}")
    print(f"🎉 Save successful!")
    print(f"📄 File: {OUTPUT_FILE}")
    print(f"📋 Number of rows: {len(all_results)}")
    print(f"🏢 Number of companies: {len(csv_files)}")


# ============================================================
# รันฟังก์ชัน
# ============================================================
if __name__ == "__main__":
    main()