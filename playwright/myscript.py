# ============================================================
# Playwright Script: ดึงข้อมูล Meeting Attendance จาก SET Smart
# ============================================================
# วิธีติดตั้ง (ครั้งแรกครั้งเดียว):
#   1. pip install playwright
#   2. playwright install
#
# วิธีรัน:
#   python setsmart_login.py
#
# Output:
#   โฟลเดอร์ output/ จะมีไฟล์ CSV แยกตาม symbol เช่น:
#     output/2S_meeting_attendance.csv
#     output/3K-BAT_meeting_attendance.csv
#     output/A_meeting_attendance.csv
# ============================================================

from playwright.sync_api import sync_playwright
import time
import csv
import os

# ============================================================
# ⚠️ ใส่ username และ password ของคุณตรงนี้
# ============================================================
USERNAME = "smart-research2"
PASSWORD = "Rese@rt#2"

# ============================================================
# รายชื่อ Symbol ที่ต้องการ
# ============================================================
SYMBOLS = ["F%26D", "L%26E", "S%26J"]

    # 'FM', 'SNPS', 'TERA', 'CHAO', 'MAGURO', 'MRDIYT', 'PCE', 'SEI', 'HANN', 
    # 'GVREIT', 'LTS', 'MEDEZE', 'AIMIRT', 'TATG', 'IROYAL', 'MOTHER', 'STX', 
    # 'SPREME', 'UNIX', 'BKA', 'NKT', 'GULF', 'BKIH', '2023', 'ATLAS', 'CPAXT', 
    # 'OKJ', 'WASH', 'MMM', 'AIMCG', 'QTCG', 'NTF', 'CFARM', 'ONSENS', 'NEO', 
    # 'NUT', 'TURBO', 'SMO', 'TMAN', 'STECON', 'LTMH', 'MASTEC', 'KCC', 'PIS', 
    # 'PMC', 'APO', '88TH', 'FTREIT', 'NCP', 'SKIN', 'IVF', 'MPJ', 'IDG']

    # ["2S","3K-BAT","A","A5","AAI","AAV","ABM","ACAP","ACC","ACE",
    # "ACG","ADB","ADD","ADVANC","ADVICE","AE","AEONTS","AF","AFC","AGE",
    # "AH","AHC","AI","AIE","AIRA","AIT","AJ","AJA","AKP","AKR",
    # "AKS","ALL","ALLA","ALLY","ALPHAX","ALT","ALUCON","AMA","AMANAH","AMARC",
    # "AMARIN","AMATA","AMATAV","AMC","AMR","ANAN","ANI","AOT","AP","APCO",
    # "APCS","APP","APURE","AQUA","ARIN","ARIP","ARROW","AS","ASAP","ASEFA",
    # "ASIA","ASIAN","ASIMAR","ASK","ASN","ASP","ASW","ATP30","AU","AUCT",
    # "AURA","AWC","AYUD","B","B52","BA","BAFS","BAM","BANPU","BAY",
    # "BBGI","BBIK","BBL","BC","BCH","BCP","BCPG","BCT","BDMS","BE8",
    # "BEAUTY","BEC","BEM","BEYOND","BGC","BGRIM","BGT","BH","BIG","BIOTEC",
    # "BIS","BIZ","BJC","BJCHI","BKD","BKGI","BKI","BLA","BLAND","BLC",
    # "BLESS","BLISS","BM","BOL","BPP","BPS","BR","BRI","BROCK","BRR",
    # "BSBM","BSM","BSRC","BTC","BTG","BTNC","BTS","BTW","BUI","BVG",
    # "BWG","BYD","CAZ","CBG","CCET","CCP","CEN","CENTEL","CEYE","CFRESH",
    # "CGD","CGH","CH","CHARAN","CHASE","CHAYO","CHEWA","CHG","CHIC","CHO",
    # "CHOTI","CHOW","CI","CIG","CIMBT","CITY","CIVIL","CK","CKP","CM",
    # "CMAN","CMC","CMO","CMR","CNT","COCOCO","COLOR","COM7","COMAN","COTTO",
    # "CPALL","CPANEL","CPAXTT","CPF","CPH","CPI","CPL","CPN","CPNREIT","CPR",
    # "CPT","CPW","CRANE","CRC","CRD","CREDIT","CSC","CSP","CSR","CSS",
    # "CTW","CV","CWT","D","DCC","DCON","DDD","DELTA","DEMCO","DEXON",
    # "DHOUSE","DIMET","DITTO","DMT","DOD","DOHOME","DPAINT","DREIT","DRT","DTCENT",
    # "DTCI","DUSIT","DV8","EA","EASON","EAST","EASTW","ECF","EFORL","EGCO",
    # "EKH","EMC","EMPIRE","EP","EPG","ERW","ESTAR","ETC","ETE","ETL",
    # "EURO","EVER","F&D","FANCY","FE","FLOYD","FMT","FN","FNS","FORTH",
    # "FPI","FPT","FSMART","FSS","FSX","FTE","FTI","FVC","GABLE","GBX",
    # "GC","GCAP","GEL","GENCO","GFC","GFPT","GGC","GJS","GLAND","GLOBAL",
    # "GLORY","GPI","GPSC","GRAMMY","GRAND","GREEN","GSTEEL","GTB","GTV","GULFI",
    # "GUNKUL","GYT","HANA","HARN","HEALTH","HENG","HFT","HL","HMPRO","HPT",
    # "HTC","HTECH","HUMAN","HYDRO","I2","ICC","ICHI","ICN","IFEC","IFS",
    # "IHL","IIG","III","ILINK","ILM","IMH","IMPACT","IND","INET","INGRS",
    # "INOX","INSET","INSURE","INTUCH","IP","IRC","IRCP","IRPC","IT","ITC",
    # "ITD","ITEL","ITNS","ITTHI","IVL","J","JAK","JAS","JCK","JCT",
    # "JDF","JKN","JMART","JMT","JPARK","JR","JSP","JTS","JUBILE","K",
    # "KAMART","KASET","KBANK","KBS","KC","KCAR","KCCAMC","KCE","KCG","KCM",
    # "KDH","KEX","KGEN","KGI","KIAT","KISS","KJL","KK","KKC","KKP",
    # "KLINIQ","KOOL","KSL","KTB","KTC","KTIS","KTMS","KUMWEL","KUN","KWC",
    # "KWI","KWM","KYE","L&E","LALIN","LANNA","LDC","LEE","LEO","LH",
    # "LHFG","LHK","LIT","LOXLEY","LPH","LPN","LRH","LST","M","MADAME",
    # "MAJOR","MALEE","MANRIN","MASTER","MATCH","MATI","MBAX","MBK","MC","MCA",
    # "M-CHAI","MCOT","MCS","MDX","MEB","MEGA","MENA","META","METCO","MFC",
    # "MFEC","MGC","MGI","MGT","MICRO","MIDA","MILL","MINT","MITSIB","MJD",
    # "MK","ML","MODERN","MONO","MOONG","MORE","MOSHI","MSC","MST","MTC",
    # "MTI","MTW","MUD","MVP","NAM","NAT","NATION","NC","NCAP","NCH",
    # "NCL","NDR","NEP","NER","NETBAY","NEW","NEX","NFC","NKI","NL",
    # "NNCL","NOBLE","NOK","NOVA","NPK","NRF","NSI","NSL","NTL","NTSC",
    # "NTV","NV","NVD","NWR","NYT","OCC","OGC","OHTL","ONEE","OR",
    # "ORI","ORN","OSP","PACO","PAF","PANEL","PAP","PATO","PB","PCC",
    # "PCSGH","PDG","PDJ","PEACE","PEER","PERM","PF","PG","PHG","PHOL",
    # "PICO","PIMO","PIN","PJW","PK","PL","PLANB","PLANET","PLAT","PLE",
    # "PLT","PLUS","PM","PMTA","POLAR","POLY","POMPUI","PORT","POST","PPM",
    # "PPP","PPPM","PPS","PQS","PR9","PRAKIT","PRAPAT","PREB","PRECHA","PRG",
    # "PRI","PRIME","PRIN","PRINC","PRM","PRO","PROEN","PROS","PROUD","PRTR",
    # "PSGC","PSH","PSL","PSP","PSTC","PT","PTC","PTECH","PTG","PTL",
    # "PTT","PTTEP","PTTGC","PYLON","Q-CON","QDC","QH","QLT","QTC","RABBIT",
    # "RAM","RATCH","RBF","RCL","READY","RICHY","RJH","RML","ROCK","ROCTEC",
    # "ROH","ROJNA","RP","RPC","RPH","RS","RSP","RT","RWI","S",
    # "S&J","S11","SA","SAAM","SABINA","SAF","SAFARI","SAFE","SAK","SALEE",
    # "SAM","SAMART","SAMCO","SAMTEL","SANKO","SAPPE","SAT","SAUCE","SAV","SAWAD",
    # "SAWANG","SC","SCAP","SCB","SCC","SCCC","SCG","SCGD","SCGP","SCI",
    # "SCL","SCM","SCN","SCP","SDC","SE","SEAFCO","SEAOIL","SECURE","SE-ED",
    # "SELIC","SENA","SENX","SFLEX","SFT","SGC","SGF","SGP","SHANG","SHR",
    # "SIAM","SICT","SIMAT","SINGER","SINO","SIRI","SIS","SISB","SITHAI","SJWD",
    # "SK","SKE","SKN","SKR","SKY","SLM","SLP","SMART","SMD100","SMIT",
    # "SMK","SMPC","SMT","SNC","SNNP","SNP","SO","SOLAR","SONIC","SORKON",
    # "SPA","SPACK","SPALI","SPC","SPCG","SPG","SPI","SPRC","SPTX","SPVI",
    # "SQ","SR","SRICHA","SRS","SSC","SSF","SSP","SSSC","SST","STA",
    # "STANLY","STARK","STARM","STC","STEC","STECH","STELLA","STGT","STHAI","STI",
    # "STOWER","STP","STPI","SUC","SUN","SUPER","SUSCO","SUTHA","SVI","SVOA",
    # "SVR","SVT","SWC","SYMC","SYNEX","SYNTEC","TACC","TAE","TAKUNI","TAN",
    # "TAPAC","TASCO","TBN","TC","TCAP","TCC","TCCC","TCJ","TCMC","TCOAT",
    # "TEAM","TEAMG","TEGH","TEKA","TFG","TFI","TFM","TFMAMA","TGE","TGH",
    # "TGPRO","TH","THAI","THANA","THANI","THCOM","THE","THG","THIP","THL",
    # "THMUI","THRE","THREL","TIGER","TIPCO","TIPH","TISCO","TITLE","TK","TKC",
    # "TKN","TKS","TKT","TL","TLI","TM","TMC","TMD","TMI","TMILL",
    # "TMT","TMW","TNDT","TNH","TNITY","TNL","TNP","TNPC","TNR","TOA",
    # "TOG","TOP","TOPP","TPA","TPAC","TPBI","TPCH","TPCS","TPIPL","TPIPP",
    # "TPL","TPLAS","TPOLY","TPP","TPS","TQM","TQR","TR","TRC","TRITN",
    # "TRP","TRT","TRU","TRUBB","TRUE","TRV","TSC","TSE","TSI","TSR",
    # "TSTE","TSTH","TTA","TTB","TTCL","TTI","TTT","TTW","TU","TURTLE",
    # "TVDH","TVH","TVI","TVO","TVT","TWP","TWPC","TWZ","TYCN","UAC",
    # "UBA","UBE","UBIS","UEC","UKEM","UMI","UMS","UNIQ","UOBKH","UP",
    # "UPF","UPOIC","UREKA","UTP","UV","UVAN","VARO","VCOM","VGI","VIBHA",
    # "VIH","VL","VNG","VPO","VRANDA","VS","WACOAL","WARRIX","WAVE","WELL",
    # "WFX","WGE","WHA","WHAUP","WICE","WIIK","WIN","WINDOW","WINMED","WINNER",
    # "WORK","WORLD","WP","WPH","WSOL","XBIO","XO","XPG","XYZ","YGG",
    # "YONG","YUASA","ZAA","ZEN","ZIGA"]

# ปีที่ต้องการดึงข้อมูล
YEARS = [2021, 2022, 2023, 2024, 2025, 2026]

# โฟลเดอร์เก็บ output
OUTPUT_DIR = "output"


def select_dropdown(page, dropdown_id, value_text, parent_selector=None):
    """
    ฟังก์ชันช่วย: กดเปิด dropdown แล้วเลือกค่าที่ต้องการ

    Parameters:
      page             = หน้าเว็บ Playwright
      dropdown_id      = id ของปุ่ม dropdown เช่น "dropdown-begin-year"
      value_text       = ข้อความที่ต้องการเลือก เช่น "2021"
      parent_selector  = (ถ้ามี) CSS selector ของ parent element
                         ใช้เมื่อหน้าเว็บมี id ซ้ำกันหลายตัว
                         เช่น "app-meeting-attendance" จะหา dropdown
                         เฉพาะที่อยู่ภายใน <app-meeting-attendance>

    ปัญหาที่เจอ:
      หน้า governance มี #dropdown-year ซ้ำกัน 8 ตัว!
      (แต่ละ section เช่น Board of Director, Audit Committee ฯลฯ
       มี dropdown ของตัวเอง แต่ใช้ id เดียวกันหมด)

      ถ้าใช้ page.locator("#dropdown-year") จะ error เพราะ
      Playwright เจอ 8 ตัว ไม่รู้จะกดตัวไหน (strict mode violation)

    วิธีแก้:
      ใส่ parent_selector เช่น "app-meeting-attendance"
      จะได้ selector เป็น "app-meeting-attendance #dropdown-year"
      ซึ่งหาได้ตัวเดียวที่อยู่ใน section Meeting Attendance
    """
    # สร้าง selector สำหรับปุ่ม dropdown
    # ถ้ามี parent_selector → "parent #dropdown-id"
    # ถ้าไม่มี            → "#dropdown-id"
    if parent_selector:
        btn_selector = f"{parent_selector} #{dropdown_id}"
    else:
        btn_selector = f"#{dropdown_id}"

    # กดเปิด dropdown (ใช้ JavaScript เพราะอาจมี element ทับ)
    page.locator(btn_selector).first.evaluate("el => el.click()")
    time.sleep(1)

    # ---- วิธีที่ 1: หาจาก id="selected-XXXX" ใน dropdown menu ที่เปิดอยู่ ----
    item_selector = f".dropdown-menu.show #selected-{value_text}"
    item = page.locator(item_selector)

    if item.count() > 0:
        item.first.evaluate("el => el.click()")
        time.sleep(0.5)
        return

    # ---- วิธีที่ 2: หาจากข้อความใน dropdown menu ที่เปิดอยู่ (.show) ----
    item2 = page.locator(f".dropdown-menu.show li").filter(has_text=value_text).first
    if item2.count() > 0:
        item2.evaluate("el => el.click()")
        time.sleep(0.5)
        return

    # ---- วิธีที่ 3 (fallback): หาจากทุก dropdown-menu ----
    page.locator(f".dropdown-menu >> text='{value_text}'").first.evaluate("el => el.click()")
    time.sleep(0.5)


def extract_summary_info(page):
    """
    ดึงข้อมูลสรุป 4 รายการจาก meeting-attendance-table:
      - Number of the Board of Directors meeting
      - Number of the audit committee meeting
      - Date of AGM meeting
      - EGM meeting

    Returns:
      dict เช่น {
        "Board Meeting Count": "5",
        "Audit Committee Meeting Count": "4",
        "AGM Date": "24/04/2024",
        "EGM": "✓" หรือ "✗"
      }
    """
    container = page.locator("#meeting-attendance-table .margin-card-container > .col-lg-6, #meeting-attendance-table > .col-lg-6")

    # ถ้าหา container ไม่เจอ ลองหาแบบกว้างกว่า
    if container.count() == 0:
        container = page.locator("#meeting-attendance-table")

    summary = {}

    # ดึงทุก row ที่เป็น .row ภายใน container
    rows = page.locator("#meeting-attendance-table .row")
    for idx in range(rows.count()):
        row = rows.nth(idx)
        text = row.text_content().strip()

        if "Board of Directors meeting" in text and "Number" in text:
            # หาตัวเลข: อยู่ใน .col-2
            val = row.locator(".col-2").first.text_content().strip()
            summary["Board Meeting Count"] = val
        elif "audit committee meeting" in text and "Number" in text:
            val = row.locator(".col-2").first.text_content().strip()
            summary["Audit Committee Meeting Count"] = val
        elif "Date of AGM" in text:
            val = row.locator(".col-5").first.text_content().strip()
            summary["AGM Date"] = val
        elif "EGM meeting" in text:
            # EGM อาจเป็น ✓ (fa-check) หรือ ✗ (fa-times) หรือข้อความ
            val = row.locator(".col-5").first.text_content().strip()
            if not val:
                # ตรวจสอบ icon
                has_check = row.locator(".fa-check").count() > 0
                has_times = row.locator(".fa-times").count() > 0
                if has_check:
                    val = "Yes"
                elif has_times:
                    val = "No"
                else:
                    val = "N/A"
            summary["EGM"] = val

    return summary


def extract_attendance_table(page):
    """
    ดึงข้อมูลตาราง Meeting Attendance (Board of Directors + Audit Committee + อื่นๆ)

    Returns:
      list ของ dict เช่น:
      [
        {"Section": "Board of Directors", "No": "1", "Name": "KUNCHIT...",
         "Termination Date": "N/A", "Number of Board Meeting": "4 / 4",
         "%": "100.00", "AGM Meetings": "Yes", "EGM Meetings": ""},
        ...
      ]

    หมายเหตุ:
      - แถว Board of Directors มี 7 คอลัมน์ (รวม AGM, EGM)
      - แถว Audit Committee มีแค่ 5 คอลัมน์ (ไม่มี AGM, EGM)
      - แถว อื่นๆ อาจมี 5 หรือ 7 คอลัมน์
      - AGM/EGM อาจเป็น icon (fa-check / fa-times) ไม่ใช่ข้อความ
    """
    rows_data = []

    # หาตาราง meeting-attendance
    table = page.locator("#meeting-attendance-table table.meeting-attendance")

    if table.count() == 0:
        return rows_data

    # ดึงทุกแถวใน tbody
    tbody_rows = table.locator("tbody tr")
    current_section = ""

    for idx in range(tbody_rows.count()):
        row = tbody_rows.nth(idx)
        cells = row.locator("td")
        cell_count = cells.count()

        if cell_count == 0:
            continue

        # เช็คว่าเป็น header row (colspan = 100%)
        first_cell = cells.first
        colspan = first_cell.get_attribute("colspan")

        if colspan:
            text = first_cell.text_content().strip()

            if "No Information" in text:
                rows_data.append({
                    "Section": current_section,
                    "No": "",
                    "Name": "No Information",
                    "Termination Date": "",
                    "Number of Board Meeting": "",
                    "%": "",
                    "AGM Meetings": "",
                    "EGM Meetings": "",
                })
            else:
                current_section = text
            continue

        # ----- แถวข้อมูลปกติ -----
        # คอลัมน์ 0-4 มีเสมอ: No, Name, Termination Date, Board Meeting, %
        if cell_count >= 5:
            row_dict = {
                "Section": current_section,
                "No": cells.nth(0).text_content().strip(),
                "Name": cells.nth(1).text_content().strip(),
                "Termination Date": cells.nth(2).text_content().strip(),
                "Number of Board Meeting": cells.nth(3).text_content().strip(),
                "%": cells.nth(4).text_content().strip(),
                "AGM Meetings": "",
                "EGM Meetings": "",
            }

            # คอลัมน์ 5 = AGM (ถ้ามี)
            # ค่าอาจเป็น:
            #   - icon fa-check → "Yes"
            #   - icon fa-times → "No"
            #   - ข้อความ → ใช้ข้อความนั้น
            #   - ว่าง → ""
            if cell_count >= 6:
                agm_cell = cells.nth(5)
                row_dict["AGM Meetings"] = _read_icon_or_text(agm_cell)

            # คอลัมน์ 6 = EGM (ถ้ามี)
            if cell_count >= 7:
                egm_cell = cells.nth(6)
                row_dict["EGM Meetings"] = _read_icon_or_text(egm_cell)

            rows_data.append(row_dict)

    return rows_data


def _read_icon_or_text(cell):
    """
    อ่านค่าจาก cell ที่อาจเป็น icon หรือข้อความ

    เว็บนี้ใช้:
      <i class="fas fa-check">  = เข้าร่วม (✓)
      <i class="fas fa-times">  = ไม่เข้าร่วม (✗)
      ข้อความธรรมดา             = ใช้ข้อความนั้น
      ว่าง                       = ""

    Returns:
      "Yes", "No", ข้อความ, หรือ ""
    """
    # เช็ค icon ก่อน
    has_check = cell.locator("i.fa-check").count() > 0
    has_times = cell.locator("i.fa-times").count() > 0

    if has_check:
        return "Yes"
    elif has_times:
        return "No"
    else:
        text = cell.text_content().strip()
        return text if text else ""


def login_and_scrape():
    """
    ฟังก์ชันหลัก:
      1. Login
      2. วนแต่ละ symbol → เข้า governance page
      3. ตั้ง begin year / end year
      4. วน dropdown year → ดึงข้อมูล Meeting Attendance
      5. บันทึก CSV แยกไฟล์ต่อ symbol
    """

    # สร้างโฟลเดอร์ output
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with sync_playwright() as p:

        # ===========================================
        # ขั้นตอนที่ 1: เปิด Browser + Login
        # ===========================================
        browser = p.chromium.launch(
            headless=False,
            slow_mo=300
        )
        page = browser.new_page()

        print("🌐 กำลังเปิดหน้า Login ...")
        page.goto("https://www.setsmart.com/ssm/login")
        page.wait_for_selector("#username", timeout=15000)

        print("📝 กรอก Username + Password ...")
        page.fill("#username", USERNAME)
        page.fill("#password", PASSWORD)
        page.click("#login-btn")

        print("⏳ รอ Login ...")
        try:
            page.wait_for_url("**/marketSummary", timeout=15000)
            print("🎉 Login สำเร็จ!\n")
        except Exception:
            print("❌ Login unsucceddful")
            input("🛑 Click Enter for close ...")
            browser.close()
            return

        # ===========================================
        # ขั้นตอนที่ 2: วนแต่ละ Symbol
        # ===========================================
        for sym_idx, symbol in enumerate(SYMBOLS, start=1):

            print(f"\n{'='*60}")
            print(f"📊 [{sym_idx}/{len(SYMBOLS)}] Symbol: {symbol}")
            print(f"{'='*60}")

            # ข้อมูลทุกปีของ symbol นี้จะเก็บไว้ที่นี่
            all_year_data = []

            # -------------------------------------------
            # ขั้นตอนที่ 3: เข้าหน้า Governance ตรงๆ
            # -------------------------------------------
            # ไม่ต้องเข้า highlight ก่อน → ไป governance เลย
            gov_url = f"https://www.setsmart.com/ssm/esg-data/governance?symbol={symbol}"
            print(f"   🌐 Entering Governance: {gov_url}")
            page.goto(gov_url)
            # ใช้ domcontentloaded แทน networkidle
            # เพราะเว็บ Angular มักมี request ส่งอยู่ตลอด ทำให้ networkidle timeout
            # domcontentloaded = รอแค่ HTML โหลดเสร็จ แล้วใช้ sleep รอข้อมูลเพิ่ม
            page.wait_for_load_state("domcontentloaded")
            time.sleep(5)  # รอ Angular render ข้อมูลให้เสร็จ

            # -------------------------------------------
            # ขั้นตอนที่ 4: เช็คว่าหน้ามีข้อมูลไหม
            # -------------------------------------------
            # บริษัทบางตัวออกจากตลาดแล้ว หน้าอาจไม่มี dropdown
            # ถ้าหา dropdown-begin-year ไม่เจอภายใน 5 วิ → ข้ามไป symbol ถัดไป
            if page.locator("#dropdown-begin-year").count() == 0:
                print(f"   ⏭️ skip {symbol} — No data (may have delisted)")
                continue

            # -------------------------------------------
            # ขั้นตอนที่ 5: ตั้ง Begin Year = 2021 และ End Year = 2026
            # -------------------------------------------
            # โครงสร้าง HTML ของ dropdown:
            #   <button id="dropdown-begin-year">...</button>
            #   <ul class="dropdown-menu">
            #     <li><a id="item-2026" class="dropdown-item">2026</a></li>
            #     <li><a id="item-2025" class="dropdown-item">2025</a></li>
            #     ...
            #     <li><a id="item-2021" class="dropdown-item">2021</a></li>
            #   </ul>
            #
            # วิธีทำ: กดเปิด dropdown → คลิก #item-YYYY ตรงๆ

            for dd_id, target_year in [
                ("dropdown-begin-year", str(YEARS[0])),
                ("dropdown-end-year", str(YEARS[-1])),
            ]:
                label = "Begin Year" if "begin" in dd_id else "End Year"
                print(f"   📅 Setting {label} = {target_year} ...")
                try:
                    # กดเปิด dropdown
                    page.locator(f"#{dd_id}").first.evaluate("el => el.click()")
                    time.sleep(1)

                    # ⚠️ ปัญหา: #item-2021 อาจซ้ำกันหลายตัวในหน้า
                    # (begin year, end year, และ dropdown-year อื่นๆ ใช้ id เดียวกัน)
                    #
                    # วิธีแก้: คลิกเฉพาะ #item-YYYY ที่อยู่ใน .dropdown-menu.show
                    # (menu ที่เปิดอยู่ตอนนี้)
                    item = page.locator(f".dropdown-menu.show #item-{target_year}")

                    if item.count() > 0:
                        item.first.evaluate("el => el.click()")
                        print(f"   ✅ Successfully set {label} = {target_year}")
                    else:
                        print(f"   ⚠️ Cannot find {target_year} in dropdown")

                    time.sleep(1)
                except Exception as e:
                    print(f"   ⚠️ Cannot set {label}: {e}")

            # -------------------------------------------
            # ขั้นตอนที่ 6: เลื่อนลงไปหา Meeting Attendance
            # -------------------------------------------
            print("   📜 Scrolling to Meeting Attendance ...")
            try:
                # หา header "Meeting Attendance of the Board..."
                header = page.locator("div.header-sub-title", has_text="Meeting Attendance")
                header.scroll_into_view_if_needed()
                time.sleep(1)
            except Exception:
                print("   ⚠️ Cannot find Meeting Attendance section")

            # -------------------------------------------
            # ขั้นตอนที่ 7: วนเลือก dropdown-year ทีละปี
            # -------------------------------------------
            # dropdown-year คือ dropdown ที่อยู่ใกล้ตาราง Meeting Attendance
            # เลือกปีแล้วข้อมูลในตารางจะเปลี่ยน

            for year in YEARS:
                print(f"\n   📆 year {year} ...")

                try:
                    # กดเปิด dropdown-year ใน section meeting-attendance
                    # ใช้ parent "app-meeting-attendance" เพราะ #dropdown-year ซ้ำ 8 ตัว
                    page.locator("app-meeting-attendance #dropdown-year").first.evaluate("el => el.click()")
                    time.sleep(1)

                    # คลิก item เฉพาะที่อยู่ใน .dropdown-menu.show
                    item = page.locator(f".dropdown-menu.show #item-{year}")
                    if item.count() > 0:
                        item.first.evaluate("el => el.click()")
                        print(f"      ✅ เลือกปี {year} สำเร็จ")
                    else:
                        print(f"      ⚠️ ไม่เจอปี {year} ใน dropdown")

                    time.sleep(2)  # รอข้อมูลโหลด
                except Exception as e:
                    print(f"      ⚠️ เลือกปี {year} ไม่ได้: {e}")
                    # เพิ่มแถวว่างสำหรับปีนี้
                    all_year_data.append({
                        "Symbol": symbol,
                        "Year": year,
                        "Board Meeting Count": "N/A",
                        "Audit Committee Meeting Count": "N/A",
                        "AGM Date": "N/A",
                        "EGM": "N/A",
                        "Section": "",
                        "No": "",
                        "Name": f"Error selecting year: {e}",
                        "Termination Date": "",
                        "Number of Board Meeting": "",
                        "%": "",
                        "AGM Meetings": "",
                        "EGM Meetings": "",
                    })
                    continue

                # ดึงข้อมูลสรุป (จำนวนครั้งประชุม, วัน AGM, EGM)
                summary = extract_summary_info(page)
                print(f"      สรุป: {summary}")

                # ดึงตารางรายชื่อ attendance
                table_rows = extract_attendance_table(page)
                print(f"      ตาราง: {len(table_rows)} แถว")

                if table_rows:
                    for row in table_rows:
                        row["Symbol"] = symbol
                        row["Year"] = year
                        row["Board Meeting Count"] = summary.get("Board Meeting Count", "")
                        row["Audit Committee Meeting Count"] = summary.get("Audit Committee Meeting Count", "")
                        row["AGM Date"] = summary.get("AGM Date", "")
                        row["EGM"] = summary.get("EGM", "")
                        all_year_data.append(row)
                else:
                    # ไม่มีข้อมูลในตาราง → เก็บแถวสรุปอย่างน้อย
                    all_year_data.append({
                        "Symbol": symbol,
                        "Year": year,
                        "Board Meeting Count": summary.get("Board Meeting Count", "N/A"),
                        "Audit Committee Meeting Count": summary.get("Audit Committee Meeting Count", "N/A"),
                        "AGM Date": summary.get("AGM Date", "N/A"),
                        "EGM": summary.get("EGM", "N/A"),
                        "Section": "",
                        "No": "",
                        "Name": "No Data",
                        "Termination Date": "",
                        "Number of Board Meeting": "",
                        "%": "",
                        "AGM Meetings": "",
                        "EGM Meetings": "",
                    })

            # -------------------------------------------
            # ขั้นตอนที่ 8: บันทึก CSV สำหรับ Symbol นี้
            # -------------------------------------------
            csv_filename = os.path.join(OUTPUT_DIR, f"{symbol}_meeting_attendance.csv")

            # คอลัมน์ทั้งหมดที่จะบันทึก
            fieldnames = [
                "Symbol", "Year",
                "Board Meeting Count", "Audit Committee Meeting Count",
                "AGM Date", "EGM",
                "Section", "No", "Name", "Termination Date",
                "Number of Board Meeting", "%",
                "AGM Meetings", "EGM Meetings",
            ]

            # เขียนไฟล์ CSV
            # encoding="utf-8-sig" = ใส่ BOM เพื่อให้ Excel เปิดภาษาไทยได้ถูกต้อง
            with open(csv_filename, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()      # เขียนหัวคอลัมน์
                writer.writerows(all_year_data)  # เขียนข้อมูลทุกแถว

            print(f"\n   💾 saved: {csv_filename}")
            print(f"   📋 Number of rows: {len(all_year_data)}")

        # ===========================================
        # ขั้นตอนที่ 9: เสร็จ
        # ===========================================
        print(f"\n{'='*60}")
        print(f"🎉 Success all Symbol! ({len(SYMBOLS)})")
        print(f"📁 Files in folder: {OUTPUT_DIR}/")
        input("\n🛑 Click Enter for close browser ...")
        browser.close()


# ============================================================
# รันฟังก์ชัน
# ============================================================
if __name__ == "__main__":
    login_and_scrape()