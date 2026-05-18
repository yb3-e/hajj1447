"""
═══════════════════════════════════════════════════════════════
  لوحة التحكم الميدانية الموحدة – حج 1447
  تطوير: م. عبدالعزيز الشهري

  المصادر:
  ① نموذج_احصاء_الموظفين_.xlsx  ← خريطة حقيقية (دبابيس ورديتين)
  ② staff_data.xlsx               ← بيانات القوى العاملة، الغياب، الإحصائيات
═══════════════════════════════════════════════════════════════
"""

import pandas as pd
import os, json, webbrowser, time, requests
from datetime import datetime, timedelta

# ============================================================
#  الإعدادات
# ============================================================
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
SURVEY_PATH  = os.path.join(BASE_DIR, "نموذج_احصاء_الموظفين_.xlsx")
STAFF_PATH   = os.path.join(BASE_DIR, "staff_data.xlsx")
REPORT_PATH  = os.path.join(BASE_DIR, "index.html")

USERNAME = 'E1126415635'
PASSWORD = '415635'

# ============================================================
#  الإحداثيات الجغرافية لكل موقع (من tanaqol_map_final)
# ============================================================
LOCATIONS_COORDS = {
    'اجياد':              [21.4192, 39.8254],
    'اجياد الداخلية':    [21.4206, 39.8256],
    'الشبيكة':           [21.4211, 39.8211],
    'الشبيكة الداخلية':  [21.4219, 39.8224],
    'مطاف السطح':        [21.4219, 39.8267],
    'مطاف الميزانين':    [21.4229, 39.8257],
    'مسعى السطح':        [21.4240, 39.8281],
    'مسعى ميزانين':      [21.4230, 39.8276],
    'الساحات الشرقية':   [21.4242, 39.8294],
    'المروة':             [21.4261, 39.8276],
    'المروة الداخلية':   [21.4253, 39.8274],
    'القشاشية':          [21.4200, 39.8245],
}

# ============================================================
#  الهيكل الإداري للورديات
# ============================================================
MANAGEMENT = {
    'الوردية الأولى': {
        'Chief_HR':         'فهد محمد السبيعي (رئيس الأخصائيين)',
        'Field_Leader':     'أحمد المحمادي (رئيس وردية ميداني)',
        'Assistant_Leader': 'فيصل ابراهيم الفهمي (مساعد مشرف الوردية)',
        'HR_Assignment': {
            'اجياد':             'مهنا صالح المالكي',
            'اجياد الداخلية':   'مهنا صالح المالكي',
            'الشبيكة':          'أحمد صالح الزهراني',
            'الشبيكة الداخلية': 'محمد الحارثي',
            'مطاف السطح':       'مسعد سعود العميري',
            'مطاف الميزانين':   'مسعد سعود العميري',
            'مسعى السطح':       'مسعد سعود العميري',
            'مسعى ميزانين':     'مسعد سعود العميري',
            'المروة':            'عماد الحارثي',
            'المروة الداخلية':  'عماد الحارثي',
            'الساحات الشرقية':  'عماد الحارثي',
        }
    },
    'الوردية الثالثة': {
        'Chief_HR':         'فهد محمد السبيعي (رئيس الأخصائيين)',
        'Field_Leader':     'علي احمد الزهراني (مشرف الوردية الميداني)',
        'Assistant_Leader': 'سمير محمد الحارثي (مساعد مشرف الوردية)',
        'HR_Assignment': {
            'اجياد':             'عبدالعزيز محمد الريشي',
            'اجياد الداخلية':   'عبدالعزيز محمد الريشي',
            'الشبيكة':          'محمد سعيد المالكي',
            'الشبيكة الداخلية': 'سراج عمر نوري',
            'مطاف السطح':       'محمد رزق الله الثبيتي',
            'مطاف الميزانين':   'محمد رزق الله الثبيتي',
            'مسعى السطح':       'محمد رزق الله الثبيتي',
            'مسعى ميزانين':     'محمد رزق الله الثبيتي',
            'المروة':            'ريان عدنان الغامدي',
            'المروة الداخلية':  'ريان عدنان الغامدي',
            'الساحات الشرقية':  'ريان عدنان الغامدي',
        }
    }
}

# ============================================================
#  دوال مساعدة
# ============================================================
def normalize_shift(s):
    """توحيد اسم الوردية"""
    s = str(s).replace('ة','ه').replace('أ','ا').replace('إ','ا').strip()
    if 'اولى' in s or 'أولى' in s or '1' in s: return 'الوردية الأولى'
    if 'ثالث' in s or '3' in s:               return 'الوردية الثالثة'
    if 'ثاني' in s or 'صباح' in s or '2' in s: return 'الوردية الثانية'
    return 'غير محدد'

def safe_extract_list(res_json):
    if not res_json: return []
    if isinstance(res_json, list): return res_json
    if isinstance(res_json, dict):
        data = res_json.get('data')
        if isinstance(data, list): return data
        if isinstance(data, dict): return data.get('list', [])
    return []

def get_hajj_token():
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.keys import Keys

        print("🌐 جاري فتح المتصفح وسحب التوكن...")
        opts = Options()
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        driver.get("https://tnql-prod.sejeltech.app/")
        time.sleep(10)
        wait = WebDriverWait(driver, 30)
        user_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text'] | (//input)[1]")))
        pass_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
        user_input.send_keys(USERNAME)
        pass_input.send_keys(PASSWORD)
        pass_input.send_keys(Keys.RETURN)
        time.sleep(15)
        token = driver.execute_script("return window.localStorage.getItem('token')") or \
                next((v for k, v in driver.execute_script("return window.localStorage;").items() if str(v).startswith("eyJ")), None)
        driver.quit()
        if token:
            print("✅ تم سحب التوكن!")
            return f"Bearer {token}"
        print("❌ ما وجدنا التوكن")
        return None
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return None

# ============================================================
#  الدالة الرئيسية
# ============================================================
def generate():
    # ── الوقت ──
    makkah_time = datetime.utcnow() + timedelta(hours=3)
    now_str     = makkah_time.strftime("%Y/%m/%d  %I:%M %p")
    today_iso   = makkah_time.strftime("%Y-%m-%d")
    hour        = makkah_time.hour

    if 8 <= hour < 16:
        shift_keywords   = ["ثاني","صباح","2"]
        active_shift_lbl = "الوردية الثانية صباح"
    elif 16 <= hour < 24:
        shift_keywords   = ["ثالث","مساء","3"]
        active_shift_lbl = "الوردية الثالثة مساء"
    else:
        shift_keywords   = ["اول","ليل","1"]
        active_shift_lbl = "الوردية الأولى ليلية"

    # ══════════════════════════════════════════════════════
    #  1. قراءة ملف الاحصاء (مصدر الخريطة الحقيقية)
    # ══════════════════════════════════════════════════════
    NAME_COL  = 'الاسم'
    JOB_COL   = 'نوع الوظيفة'
    SITE_COL  = 'اختر الموقع العام'
    MGR_COL   = 'المشرف الميداني وليس في البطاقة'
    PHONE_COL = 'الجوال'
    ID_COL    = 'رقم الهوية'

    df_survey = pd.DataFrame()
    staff_unregistered = []

    if os.path.exists(SURVEY_PATH):
        print("📂 قراءة ملف الاحصاء...")
        df_survey = pd.read_excel(SURVEY_PATH, dtype=str).fillna('غير متوفر')
        df_survey.columns = df_survey.columns.str.strip()
        df_survey['الوردية_نظيفة'] = df_survey['الوردية'].apply(normalize_shift)
        df_survey['id_clean'] = df_survey[ID_COL].str.replace('.0','', regex=False).str.strip().str.lower()
        df_survey['_source'] = 'survey'
    else:
        print("⚠️ ملف الاحصاء غير موجود، الخريطة ستكون فارغة")

    # ══════════════════════════════════════════════════════
    #  2. قراءة staff_data (للقوى العاملة + الغياب + الاحتياطي)
    # ══════════════════════════════════════════════════════
    staff_db   = {}
    excel_db   = {}
    all_excel_emps = []

    if os.path.exists(STAFF_PATH):
        print("📂 قراءة staff_data...")
        df_staff = pd.read_excel(STAFF_PATH, dtype=str).fillna('غير متوفر')
        df_staff.columns = df_staff.columns.str.strip()
        phone_col = next((c for c in ['رقم التليفون','رقم الجوال','الجوال','Phone','Mobile'] if c in df_staff.columns), None)

        for _, row in df_staff.iterrows():
            nid     = str(row.get('هويه الموظف','')).replace('.0','').strip().lower()
            name    = str(row.get('اسم الموظف','غير متوفر')).strip()
            job     = str(row.get('الوظيفه','غير متوفر')).strip()
            dept    = str(row.get('القسم','غير متوفر')).strip()
            company = str(row.get('شركة التشغيل','غير متوفر')).strip()
            shift   = str(row.get('الورديه','غير متوفر')).strip()
            etype   = str(row.get('نوع الموظف','غير متوفر')).strip()
            phone   = str(row.get(phone_col,'غير متوفر')).strip() if phone_col else 'غير متوفر'
            if nid and nid not in ('غير متوفر','nan',''):
                info = {'name':name,'job':job,'dept':dept,'company':company,
                        'phone':phone,'shift':shift,'shift_norm':normalize_shift(shift)}
                staff_db[nid] = info
                excel_db[nid] = info
            all_excel_emps.append({'name':name,'id':nid,'job':job,'company':company,
                                   'shift':shift,'type':etype,'phone':phone,'dept':dept})

        # الموظفون في staff_data وليسوا في الاحصاء → نضيفهم للخريطة
        if not df_survey.empty:
            survey_ids = set(df_survey['id_clean'].tolist())
            for nid, info in staff_db.items():
                if nid not in survey_ids:
                    staff_unregistered.append({
                        ID_COL:          nid,
                        NAME_COL:        info['name'],
                        JOB_COL:         info['job'],
                        SITE_COL:        'غير محدد',
                        MGR_COL:         'غير محدد',
                        PHONE_COL:       info['phone'],
                        'الوردية_نظيفة': info['shift_norm'],
                        'id_clean':      nid,
                        '_source':       'staff_data',
                    })
            print(f"  → {len(staff_unregistered)} موظف في staff_data غير مسجّل في الاحصاء")
    else:
        print("⚠️ ملف staff_data.xlsx غير موجود")

    # ══════════════════════════════════════════════════════
    #  3. دمج الاحصاء + غير المسجلين  → df_map (للخريطة)
    # ══════════════════════════════════════════════════════
    df_extra = pd.DataFrame(staff_unregistered) if staff_unregistered else pd.DataFrame(columns=df_survey.columns if not df_survey.empty else [])
    df_map   = pd.concat([df_survey, df_extra], ignore_index=True).fillna('غير متوفر') if not df_survey.empty else df_extra

    total_map  = len(df_map)
    total_s1   = len(df_map[df_map['الوردية_نظيفة'] == 'الوردية الأولى']) if not df_map.empty else 0
    total_s3   = len(df_map[df_map['الوردية_نظيفة'] == 'الوردية الثالثة']) if not df_map.empty else 0
    total_sites= df_map[df_map[SITE_COL] != 'غير محدد'][SITE_COL].nunique() if not df_map.empty else 0

    # ── بناء الدبابيس ──
    pins = []
    if not df_map.empty:
        grouped = df_map[df_map[SITE_COL] != 'غير محدد'].groupby(['الوردية_نظيفة', SITE_COL])
        for (shift, site), grp in grouped:
            if shift not in ('الوردية الأولى','الوردية الثالثة'):
                continue
            coords = LOCATIONS_COORDS.get(str(site).strip(), [21.4225, 39.8262])
            mgmt   = MANAGEMENT.get(shift, {})
            supervisors = sorted(set(
                str(v).strip() for v in grp[MGR_COL].unique()
                if str(v).strip() not in ('غير متوفر','لايوجد','لايوجد حالياً','nan','')
            ))
            employees = []
            for _, row in grp.iterrows():
                employees.append({
                    'name':   str(row.get(NAME_COL,'')).strip(),
                    'phone':  str(row.get(PHONE_COL,'')).replace('.0','').strip(),
                    'job':    str(row.get(JOB_COL,'')).strip(),
                    'source': row.get('_source','survey'),
                })
            pins.append({
                'shift':       shift,
                'site':        site,
                'lat':         coords[0],
                'lon':         coords[1],
                'total':       len(grp),
                'chief_hr':    mgmt.get('Chief_HR',''),
                'field_ldr':   mgmt.get('Field_Leader',''),
                'asst_ldr':    mgmt.get('Assistant_Leader',''),
                'hr_spec':     mgmt.get('HR_Assignment',{}).get(site,'مسؤول موارد بشرية ميداني'),
                'supervisors': ' ، '.join(supervisors) if supervisors else 'غير محدد',
                'employees':   employees,
            })

    # إحصائيات الورديات للنافذة المنبثقة
    stats_data = {}
    if not df_map.empty:
        for shift_name in ('الوردية الأولى','الوردية الثالثة'):
            s_df = df_map[df_map['الوردية_نظيفة'] == shift_name]
            jobs = {}
            for job, j_grp in s_df.groupby(JOB_COL):
                clean_job = str(job).strip()
                if clean_job in ('nan','غير متوفر',''): clean_job = 'غير محدد'
                jobs[clean_job] = j_grp[NAME_COL].astype(str).tolist()
            stats_data[shift_name] = {'total': len(s_df), 'jobs': jobs}

    # ══════════════════════════════════════════════════════
    #  4. API (للقوى العاملة + الغياب)
    # ══════════════════════════════════════════════════════
    token = get_hajj_token()
    if not token:
        print("❌ فشل التوكن. يتم البناء من الإكسيل فقط.")
        df = pd.DataFrame(all_excel_emps).rename(columns={
            'job':'occupationName','company':'operatorCompanyName','shift':'workShiftName'
        })
        if not df.empty:
            df['clean_name']  = df['name']
            df['clean_phone'] = df['phone']
            df['clean_dept']  = df.get('dept', 'غير متوفر')
            df['nationalId']  = df['id']
            df['employeeCode']= df['id']
            df['mapped_type'] = df['type']
        present_ids = set()
    else:
        headers = {"authorization": token, "content-type": "application/json", "lang": "ar"}
        payload = {
            "paging":{"sortField":"Id","searchOrder":2,"pageIndex":1,"totalRowsCount":10469,
                      "totalPages":1,"pageSize":11000,"sortBy":"Id Desc"},
            "data":{"searchText":"","name":"","EmployeeId":None,"OccupationIds":[],
                    "DepartmentIds":[],"SectionIds":[],"WorkShiftIds":[],"EmployeeTypes":[],
                    "ManagerIds":[],"OperatorCompanyIds":[],"NationalIdExpired":[],
                    "ActiveStatus":[True],"isPrinted":None,"isDeleted":False}
        }
        try:
            r_emp = requests.post("https://tnql-prod.sejeltech.app/api/StaffMember/GetStaffMember", headers=headers, json=payload)
            all_employees = safe_extract_list(r_emp.json())
            r_att = requests.post("https://tnql-prod.sejeltech.app/api/EmployeeAttendanceMonitor/GetAttendance", headers=headers, json=payload)
            att_data = safe_extract_list(r_att.json())
        except Exception as e:
            print(f"❌ خطأ API: {e}"); return

        present_ids = set()
        today_ar1 = makkah_time.strftime("%d/%m/%Y")
        today_ar2 = f"{makkah_time.day}/{makkah_time.month}/{makkah_time.year}"
        for x in att_data:
            if isinstance(x, dict):
                row_str = str(x.values()).lower()
                if today_iso in row_str or today_ar1 in row_str or today_ar2 in row_str:
                    if "غائب" not in row_str and "absent" not in row_str:
                        ec = x.get('employeeCode')
                        if ec: present_ids.add(str(ec).strip().lower())

        for emp in all_employees:
            nid = str(emp.get('nationalId','')).replace('.0','').strip().lower()
            api_phone = emp.get('mobileNumber') or emp.get('phoneNumber') or 'لا يوجد رقم'
            if nid in excel_db:
                ex = excel_db[nid]
                if ex['job'] not in ['غير متوفر','nan']: emp['occupationName'] = ex['job']
                if ex['company'] not in ['غير متوفر','nan']: emp['operatorCompanyName'] = ex['company']
                emp['clean_name']  = ex['name'] if ex['name'] not in ['غير متوفر','nan'] else (emp.get('name') or 'غير متوفر')
                emp['clean_dept']  = ex['dept']
                emp['clean_phone'] = ex['phone'] if ex['phone'] not in ['غير متوفر','nan'] else api_phone
            else:
                emp['clean_name']  = emp.get('name') or 'غير متوفر'
                emp['clean_dept']  = 'غير متوفر'
                emp['clean_phone'] = api_phone

        df = pd.DataFrame(all_employees).fillna('غير محدد')

        def clean_type(val):
            v = str(val).lower()
            return 'دائم' if 'permanent' in v or 'دائم' in v else ('موسمي' if 'seasonal' in v or 'موسمي' in v else 'غير محدد')
        df['mapped_type'] = df.get('employeeTypeName', pd.Series(['غير محدد']*len(df))).apply(clean_type)

    # ══════════════════════════════════════════════════════
    #  5. إحصائيات عامة (للوحة التحكم)
    # ══════════════════════════════════════════════════════
    total_employees = len(df) if not df.empty else 0
    permanent_count = len(df[df['mapped_type'] == 'دائم']) if not df.empty else 0
    seasonal_count  = len(df[df['mapped_type'] == 'موسمي']) if not df.empty else 0
    total_companies = df['operatorCompanyName'].nunique() if not df.empty else 0
    total_shifts    = df['workShiftName'].nunique() if not df.empty else 0

    # ── بناء قسم الشركات ──
    companies_html = ""
    if not df.empty:
        for company in df['operatorCompanyName'].unique():
            if pd.isna(company) or company in ('غير محدد','غير متوفر'): continue
            c_df = df[df['operatorCompanyName'] == company]
            companies_html += f"<div class='company-card'><div class='company-title'>🏢 {company}</div><div class='shift-grid'>"
            for shift in c_df['workShiftName'].unique():
                s_df = c_df[c_df['workShiftName'] == shift]
                shift_total = len(s_df)
                jobs_html = ""
                for job_name, count in s_df['occupationName'].value_counts().items():
                    j_df = s_df[s_df['occupationName'] == job_name]
                    emps_li = ""
                    for _, row in j_df.iterrows():
                        e_name  = row.get('clean_name', 'غير متوفر')
                        e_id    = str(row.get('nationalId','')).replace('.0','')
                        e_phone = row.get('clean_phone','لا يوجد رقم')
                        emps_li += f"<li><span class='emp-name'>👤 {e_name}</span><span class='emp-meta'><span>💳 {e_id}</span><span>📱 {e_phone}</span></span></li>"
                    jobs_html += f"<li><details><summary class='job-summary'><span>{job_name}</span><span class='job-val'>{count} ▾</span></summary><ul class='emp-details-list'>{emps_li}</ul></details></li>"
                companies_html += f"<div class='shift-box'><span class='shift-name'><span>📍 {shift}</span><span class='shift-total-badge'>العدد: {shift_total}</span></span><ul class='jobs-list'>{jobs_html}</ul></div>"
            companies_html += "</div></div>"

    # ── إجمالي الورديات ──
    shifts_totals_html = ""
    if not df.empty:
        for sh in df['workShiftName'].unique():
            sh_df  = df[df['workShiftName'] == sh]
            sh_tot = len(sh_df)
            comp_breakdown = ""
            for comp, grp in sh_df.groupby('operatorCompanyName'):
                if pd.isna(comp) or comp in ('غير محدد','غير متوفر'): continue
                comp_breakdown += f"<span class='sb-item'><span>{comp}</span><span class='sb-num'>{len(grp)}</span></span>"
            jobs_breakdown = ""
            for jb, cnt in sh_df['occupationName'].value_counts().items():
                j_df2 = sh_df[sh_df['occupationName'] == jb]
                names_li = "".join(f"<li>{r.get('clean_name','')}</li>" for _, r in j_df2.iterrows())
                jobs_breakdown += f"""<li>
                    <details class='job-det'>
                      <summary class='job-summary'>
                        <span>{jb}</span><span class='job-val'>{cnt} ▾</span>
                      </summary>
                      <ul class='emp-details-list names-only'>{names_li}</ul>
                    </details>
                  </li>"""
            shifts_totals_html += f"""
            <div class='shift-total-card'>
              <div class='st-head'>
                <span class='st-name'>⏱ {sh}</span>
                <span class='st-badge'>{sh_tot} موظف</span>
              </div>
              <div class='st-companies'>{comp_breakdown}</div>
              <ul class='jobs-list st-jobs'>{jobs_breakdown}</ul>
            </div>"""

    # ── ملخص الوظائف الكلي ──
    grand_jobs_html = ""
    if not df.empty:
        for j, v in df['occupationName'].value_counts().items():
            j_df3 = df[df['occupationName'] == j]
            names_li = "".join(f"<li>{r.get('clean_name','')}</li>" for _, r in j_df3.iterrows())
            grand_jobs_html += f"""<li>
              <details class='job-det'>
                <summary class='job-summary'>
                  <span>{j}</span><span class='job-val'>{v} ▾</span>
                </summary>
                <ul class='emp-details-list names-only'>{names_li}</ul>
              </details>
            </li>"""

    # ── سجل الغياب ──
    csv_data  = f"تقرير غياب: {today_iso}\nالشركة,الوردية,الاسم,الهوية,الوظيفة,القسم\n"
    absent_html = ""
    has_shift = False
    if not df.empty:
        for company in df['operatorCompanyName'].unique():
            if pd.isna(company) or company in ('غير محدد','غير متوفر'): continue
            c_df = df[df['operatorCompanyName'] == company]
            comp_abs = ""
            comp_has = False
            for shift in c_df['workShiftName'].unique():
                sh_clean = str(shift).replace('أ','ا').replace('إ','ا').replace('آ','ا').replace('ة','ه').strip().lower()
                if not any(kw in sh_clean for kw in shift_keywords): continue
                comp_has = has_shift = True
                s_df = c_df[c_df['workShiftName'] == shift]
                rows_abs = ""
                cnt_abs  = 0
                for _, row in s_df.iterrows():
                    eid = str(row.get('employeeCode','')).strip().lower()
                    nid = str(row.get('nationalId','')).replace('.0','').strip().lower()
                    if eid not in present_ids and nid not in present_ids:
                        cnt_abs += 1
                        n = row.get('clean_name',''); jo = row.get('occupationName',''); d = row.get('clean_dept','')
                        rows_abs += f"<tr><td>{n}</td><td>{nid}</td><td>{jo}</td><td>{d}</td></tr>"
                        csv_data += f"{company},{shift},{n},{nid},{jo},{d}\n"
                if cnt_abs > 0:
                    comp_abs += f"<h3 class='absent-shift-title' style='border-color:var(--danger)'>📍 {shift} <span class='abs-badge danger'>غياب: {cnt_abs}</span></h3><table class='absent-table'><tr><th>الاسم</th><th>الهوية</th><th>الوظيفة</th><th>القسم</th></tr>{rows_abs}</table>"
                else:
                    comp_abs += f"<h3 class='absent-shift-title' style='border-color:var(--success)'>📍 {shift} <span class='abs-badge success'>حضور مكتمل ✅</span></h3><div class='zero-msg'>اكتمل حضور جميع موظفي هذه الوردية.</div>"
            if comp_has:
                absent_html += f"<div class='company-card' style='margin-bottom:24px'><div class='company-title' style='font-size:1.4em'>{company}</div>{comp_abs}</div>"
    if not has_shift:
        absent_html = f"<p style='text-align:center;color:#7f8c8d;margin-top:30px'>لا توجد بيانات مسجلة لـ ({active_shift_lbl})</p>"

    # ── قائمة الأسماء الكاملة ──
    all_names_json = json.dumps([
        {"name": str(r.get('clean_name', r.get('name',''))),
         "id":   str(r.get('nationalId',r.get('id',''))).replace('.0',''),
         "job":  str(r.get('occupationName',r.get('job',''))),
         "company": str(r.get('operatorCompanyName',r.get('company',''))),
         "shift": str(r.get('workShiftName',r.get('shift','')))}
        for _, r in df.iterrows()
    ] if not df.empty else [], ensure_ascii=False)

    # JSON للخريطة
    pins_json  = json.dumps(pins,       ensure_ascii=False)
    stats_json = json.dumps(stats_data, ensure_ascii=False)

    # ══════════════════════════════════════════════════════════
    #  6. HTML الكامل الموحد
    # ══════════════════════════════════════════════════════════
    HTML = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>لوحة التحكم الميدانية – حج 1447</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<style>
:root{{
  --navy:#060F1E; --navy2:#0A1628; --navy3:#0F2040;
  --gold:#C9A84C; --gold2:#E8C97A; --gold3:#F5E3A8;
  --teal:#1AA47B; --teal2:#23C99A;
  --danger:#e74c3c; --success:#27ae60;
  --text:#D8E4F0; --text2:#7A9FC0;
  --border:rgba(201,168,76,0.18);
  --font:'IBM Plex Sans Arabic',sans-serif;
  --primary:#004d40; --secondary:#00796b; --accent:#c0a16b; --bg:#f4f7f6;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:100%;overflow:hidden}}
body{{font-family:var(--font);background:var(--navy);color:var(--text);display:flex;flex-direction:column}}
body::before{{content:'';position:fixed;inset:0;background:repeating-linear-gradient(0deg,transparent,transparent 2px,rgba(0,0,0,0.03) 2px,rgba(0,0,0,0.03) 4px);pointer-events:none;z-index:9999}}

/* ── HEADER ── */
.hdr{{flex-shrink:0;background:var(--navy2);border-bottom:1px solid var(--border);padding:0 24px;height:62px;display:flex;align-items:center;justify-content:space-between;position:relative;z-index:200}}
.hdr::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,var(--gold),transparent)}}
.hdr-brand{{display:flex;align-items:center;gap:12px}}
.hdr-logo{{width:38px;height:38px;background:linear-gradient(135deg,var(--gold),var(--gold2));border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:19px;box-shadow:0 0 14px rgba(201,168,76,0.3);flex-shrink:0}}
.hdr-title{{font-size:16px;font-weight:700;color:#fff}}
.hdr-sub{{font-size:11px;color:var(--text2);margin-top:1px}}
.hdr-center{{display:flex;gap:7px}}
.kpi{{border:1px solid var(--border);border-radius:9px;padding:6px 14px;text-align:center;min-width:80px;background:rgba(201,168,76,0.06);position:relative;overflow:hidden;transition:border-color .2s;cursor:default}}
.kpi:hover{{border-color:var(--gold)}}
.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--gold),transparent)}}
.kpi-num{{font-size:18px;font-weight:700;color:var(--gold2);display:block;line-height:1.1}}
.kpi-lbl{{font-size:10px;color:var(--text2);display:block;margin-top:1px}}
.hdr-right{{display:flex;align-items:center;gap:8px}}
.hdr-time{{font-size:11px;color:var(--text2);border:1px solid var(--border);border-radius:8px;padding:5px 11px;line-height:1.7;text-align:center}}
.hdr-time b{{display:block;color:var(--gold2);font-weight:600;font-size:12px}}
.icon-btn{{width:32px;height:32px;border-radius:7px;border:1px solid var(--border);background:rgba(201,168,76,0.07);color:var(--gold2);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:15px;transition:all .2s}}
.icon-btn:hover{{background:rgba(201,168,76,0.18);border-color:var(--gold)}}

/* ── NAV TABS ── */
.nav-tabs{{flex-shrink:0;background:var(--navy3);border-bottom:1px solid var(--border);display:flex;padding:0 24px;gap:4px;position:relative;z-index:100}}
.tab-btn{{padding:10px 20px;font-family:var(--font);font-size:13px;font-weight:600;color:var(--text2);background:transparent;border:none;border-bottom:2px solid transparent;cursor:pointer;transition:all .2s;white-space:nowrap}}
.tab-btn:hover{{color:var(--text)}}
.tab-btn.active{{color:var(--gold2);border-bottom-color:var(--gold)}}

/* ── MAIN ── */
.main{{flex:1;overflow:hidden;position:relative}}
.tab-page{{display:none;height:100%;overflow:hidden}}
.tab-page.active{{display:flex}}
.tab-page-scroll{{display:none;height:100%;overflow-y:auto}}
.tab-page-scroll.active{{display:block}}

/* ══ TAB 1: خريطة حقيقية ══ */
/* ── Shift sub-tabs ── */
.shift-subtabs{{flex-shrink:0;background:var(--navy2);border-bottom:1px solid rgba(201,168,76,.1);display:flex;padding:0 14px;gap:3px;align-items:center}}
.stab{{padding:7px 16px;font-family:var(--font);font-size:12px;font-weight:600;color:var(--text2);background:transparent;border:none;border-bottom:2px solid transparent;cursor:pointer;transition:.2s}}
.stab:hover{{color:var(--text)}}
.stab.active{{color:var(--gold2);border-bottom-color:var(--gold)}}

.map-panel{{width:280px;flex-shrink:0;background:var(--navy2);border-left:1px solid var(--border);display:flex;flex-direction:column;overflow:hidden}}
.panel-label{{padding:11px 14px 7px;font-size:10px;font-weight:700;color:var(--gold);letter-spacing:2px;text-transform:uppercase;border-bottom:1px solid rgba(255,255,255,0.05)}}
.mini-stats{{display:grid;grid-template-columns:1fr 1fr;gap:7px;padding:10px 11px}}
.mini-stat{{background:var(--navy3);border:1px solid var(--border);border-radius:8px;padding:9px 11px}}
.mini-stat .n{{font-size:17px;font-weight:700;color:var(--teal2);line-height:1}}
.mini-stat .l{{font-size:10px;color:var(--text2);margin-top:3px}}
.panel-label-wrap{{display:flex;align-items:center;justify-content:space-between;padding:10px 14px 5px}}
.site-count-badge{{font-size:10px;background:rgba(201,168,76,0.15);color:var(--gold2);padding:2px 8px;border-radius:10px;border:1px solid var(--border)}}
.search-wrap{{padding:0 10px 7px}}
.search-input{{width:100%;background:var(--navy3);border:1px solid var(--border);border-radius:8px;padding:6px 11px;color:var(--text);font-family:var(--font);font-size:12px;outline:none;transition:border-color .2s}}
.search-input::placeholder{{color:var(--text2)}}
.search-input:focus{{border-color:var(--gold)}}
.site-list{{overflow-y:auto;flex:1;padding:3px 7px 7px}}
.site-list::-webkit-scrollbar{{width:3px}}
.site-list::-webkit-scrollbar-thumb{{background:var(--gold);border-radius:2px}}
.sitem{{background:var(--navy3);border:1px solid rgba(255,255,255,0.05);border-radius:9px;padding:10px 12px;margin-bottom:5px;cursor:pointer;transition:all .2s;position:relative;overflow:hidden}}
.sitem::after{{content:'';position:absolute;right:0;top:0;bottom:0;width:2px;background:linear-gradient(180deg,var(--gold),var(--teal));opacity:0;transition:opacity .2s}}
.sitem:hover{{background:rgba(201,168,76,0.07);border-color:rgba(201,168,76,0.3)}}
.sitem:hover::after,.sitem.active::after{{opacity:1}}
.sitem.active{{background:rgba(201,168,76,0.11);border-color:rgba(201,168,76,0.5)}}
.sitem-top{{display:flex;align-items:flex-start;justify-content:space-between;gap:5px;margin-bottom:5px}}
.sitem-name{{font-size:12px;font-weight:600;color:#fff;line-height:1.35;flex:1}}
.epill{{background:rgba(26,164,123,.2);color:var(--teal2);font-size:10px;font-weight:700;padding:2px 8px;border-radius:12px;border:1px solid rgba(26,164,123,.3);flex-shrink:0}}
.sitem-shift{{font-size:10px;color:var(--gold);margin-bottom:2px}}

.map-wrap{{flex:1;position:relative}}
#map{{width:100%;height:100%}}
.leaflet-popup-content-wrapper,.leaflet-popup-tip{{background:transparent!important;box-shadow:none!important;padding:0!important;border-radius:0!important}}
.leaflet-popup-tip-container{{display:none}}
.leaflet-popup-content{{margin:0!important;width:auto!important}}

/* popup الخريطة */
.pop{{direction:rtl;font-family:var(--font);width:370px;background:#0A1628;border:1px solid #C9A84C;border-radius:13px;overflow:hidden;box-shadow:0 16px 50px rgba(0,0,0,.7);animation:popIn .18s ease}}
@keyframes popIn{{from{{opacity:0;transform:scale(.95) translateY(4px)}}to{{opacity:1;transform:none}}}}
.pop-head{{background:linear-gradient(135deg,#0C1F3F,#0F2040);padding:12px 15px;position:relative}}
.pop-head::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,#C9A84C,#23C99A,transparent)}}
.pop-title{{font-size:14px;font-weight:700;color:#fff}}
.pop-shift{{font-size:11px;color:#C9A84C;margin-top:2px}}
.pop-mgmt{{background:#f0f7ff;padding:11px 14px;border-bottom:1px solid #e2ebf4}}
.pop-row{{font-size:12px;color:#333;padding:3px 0;border-bottom:1px dashed #e6dfc7;display:flex;gap:6px;align-items:baseline}}
.pop-row:last-child{{border:none;padding-bottom:0}}
.pop-row b{{color:#1F4E78;white-space:nowrap}}
.pop-row .val{{color:#000;font-weight:600}}
.pop-row .val.teal{{color:#1AA47B}}
.pop-row .val.gold{{color:#d35400}}
.pop-emps{{padding:11px 14px 13px}}
.pop-emps-title{{font-size:10px;font-weight:700;color:#1F4E78;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:7px;display:flex;align-items:center;gap:7px}}
.pop-emps-title::after{{content:'';flex:1;height:1px;background:#e2ebf4}}
.pop-scroll{{max-height:170px;overflow-y:auto}}
.pop-scroll::-webkit-scrollbar{{width:3px}}
.pop-scroll::-webkit-scrollbar-thumb{{background:#C9A84C;border-radius:2px}}
.pop-table{{width:100%;border-collapse:collapse;font-size:11.5px}}
.pop-table th{{background:#1F4E78;color:#E8C97A;padding:5px 8px;font-weight:600;text-align:right;position:sticky;top:0}}
.pop-table td{{padding:5px 8px;color:#333;border-bottom:1px solid #f0f3f8}}
.pop-table tr:last-child td{{border:none}}
.pop-table tr:nth-child(even) td{{background:#fafbfd}}
.source-tag{{font-size:9px;background:#fff3cd;color:#856404;padding:1px 5px;border-radius:4px;margin-right:4px}}

/* الدبابيس */
.pin-wrap{{display:flex;flex-direction:column;align-items:center;cursor:pointer}}
.pin-body{{width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;font-family:var(--font);transition:all .22s;box-shadow:0 4px 14px rgba(0,0,0,.5)}}
.pin-s1{{background:#0C1F3F;border:2.5px solid #5B9BD5;color:#5B9BD5;box-shadow:0 0 0 4px rgba(91,155,213,.15),0 4px 14px rgba(0,0,0,.5)}}
.pin-s3{{background:#1A0A00;border:2.5px solid #C9A84C;color:#E8C97A;box-shadow:0 0 0 4px rgba(201,168,76,.15),0 4px 14px rgba(0,0,0,.5)}}
.pin-body.active{{transform:scale(1.22);box-shadow:0 0 0 6px rgba(255,255,255,.2),0 0 28px rgba(255,255,255,.2)}}
.pin-tail{{width:2px;height:10px;margin-top:-1px}}
.pin-s1-tail{{background:linear-gradient(180deg,#5B9BD5,transparent)}}
.pin-s3-tail{{background:linear-gradient(180deg,#C9A84C,transparent)}}

/* ══ TAB 2 & 3 ══ */
.report-page{{background:var(--bg);padding:30px;min-height:100%}}
.rp-header{{text-align:center;background:linear-gradient(135deg,var(--primary),var(--secondary));padding:40px 20px;border-radius:20px;color:white;margin-bottom:30px}}
.eng-badge{{display:inline-block;border:2px solid var(--accent);padding:8px 28px;border-radius:12px;margin-bottom:18px;background:rgba(0,0,0,0.28)}}
.eng-badge .title{{display:block;font-size:1em;color:var(--accent);font-weight:700;letter-spacing:2px}}
.eng-badge .name{{display:block;font-size:1.5em;font-weight:600}}
.rp-header h1{{font-size:2em;margin:8px 0}}
.live-indicator{{background:rgba(255,255,255,0.12);padding:6px 18px;border-radius:40px;font-weight:700;display:inline-flex;align-items:center;gap:8px;margin-top:12px;font-size:.9em}}
.pulse{{height:11px;width:11px;background:#4caf50;border-radius:50%;animation:pulse-anim 1.5s infinite;flex-shrink:0}}
@keyframes pulse-anim{{0%{{box-shadow:0 0 0 0 rgba(76,175,80,.7)}}100%{{box-shadow:0 0 0 10px rgba(76,175,80,0)}}}}
.stats-container{{display:flex;gap:12px;justify-content:center;margin:-28px 0 40px;flex-wrap:wrap;position:relative;z-index:10}}
.stat-card{{background:white;padding:18px 13px;border-radius:16px;text-align:center;box-shadow:0 8px 24px rgba(0,0,0,0.1);border-bottom:5px solid var(--accent);flex:1;min-width:130px;max-width:180px}}
.stat-card b{{display:block;font-size:2.2em;color:var(--primary);margin-bottom:4px;font-weight:700}}
.stat-card span{{font-size:.88em;font-weight:600;color:#7f8c8d}}
.company-card{{background:white;padding:28px;border-radius:22px;margin-bottom:30px;box-shadow:0 10px 30px rgba(0,0,0,0.06);border:1px solid #eee}}
.company-title{{font-size:1.7em;color:var(--primary);border-bottom:3px solid #f0f0f0;padding-bottom:15px;margin-bottom:20px;font-weight:700}}
.shift-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px}}
.shift-box{{background:#fafafa;border:1px solid #eef2f3;padding:20px;border-radius:16px}}
.shift-name{{color:var(--secondary);font-weight:700;font-size:1.15em;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;border-right:4px solid var(--accent);padding-right:12px}}
.shift-total-badge{{font-size:.72em;background:#e0f2f1;color:var(--primary);padding:4px 10px;border-radius:10px;font-weight:700}}
.jobs-list{{list-style:none;padding:0;margin:0}}
.jobs-list li{{display:block;padding:9px 0;border-bottom:1px dashed #ddd}}
.jobs-list li:last-child{{border:none}}
.job-summary{{display:flex;justify-content:space-between;align-items:center;cursor:pointer;list-style:none;outline:none;font-weight:600;color:#2c3e50}}
.job-summary::-webkit-details-marker{{display:none}}
.job-val{{background:var(--primary);color:white;padding:3px 11px;border-radius:7px;font-size:1em;font-family:monospace;font-weight:700}}
.emp-details-list{{margin-top:10px;padding:12px 14px;background:rgba(0,77,64,0.04);border-radius:9px;border-right:3px solid var(--accent);font-size:.85em;list-style:none;max-height:220px;overflow-y:auto}}
.emp-details-list li{{padding:6px 0;border-bottom:1px solid rgba(0,0,0,0.05);color:#2c3e50;font-weight:500}}
.emp-details-list li:last-child{{border:none}}
.emp-name{{color:var(--primary);font-weight:700;font-size:1.05em;display:block}}
.emp-meta{{color:#7f8c8d;display:flex;justify-content:space-between;font-size:.92em;margin-top:3px}}
.names-only li{{color:var(--primary);font-weight:600;font-size:.95em}}
.section-title{{text-align:center;color:var(--primary);font-size:1.9em;font-weight:700;margin:50px 0 25px}}
.shifts-totals-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:20px;margin-bottom:40px}}
.shift-total-card{{background:white;border:1px solid #e0e0e0;border-radius:18px;padding:22px;box-shadow:0 6px 20px rgba(0,0,0,0.05)}}
.st-head{{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;padding-bottom:10px;border-bottom:2px solid #f0f0f0}}
.st-name{{font-size:1.1em;font-weight:700;color:var(--primary)}}
.st-badge{{background:var(--primary);color:white;padding:4px 14px;border-radius:20px;font-size:.85em;font-weight:700}}
.st-companies{{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:14px}}
.sb-item{{background:#e0f2f1;border-radius:8px;padding:5px 10px;font-size:.8em;font-weight:600;color:var(--secondary);display:flex;gap:8px;align-items:center}}
.sb-num{{background:var(--secondary);color:white;border-radius:5px;padding:1px 7px;font-size:.9em}}
.st-jobs{{background:#fafafa;border-radius:10px;padding:10px 12px}}
.job-det summary{{padding:7px 0;font-size:.92em}}
.grand-summary{{background:white;border:3px solid var(--primary);padding:30px;border-radius:30px;margin-top:50px}}
.grand-summary h2{{text-align:center;color:var(--primary);font-size:1.8em;margin-bottom:25px;font-weight:700}}
.grand-jobs-list{{list-style:none;padding:0;display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:10px}}
.grand-jobs-list .job-det{{background:#f5fffe;border:1px solid #b2dfdb;border-radius:10px;padding:10px 14px}}
.grand-jobs-list .job-det summary{{font-weight:700;color:var(--primary);font-size:.95em}}
.absent-shift-title{{margin:20px 0 10px;border-right:4px solid;padding-right:10px;font-size:1.05em;display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
.abs-badge{{font-size:.78em;padding:3px 12px;border-radius:20px;color:white;font-weight:700}}
.abs-badge.danger{{background:var(--danger)}}
.abs-badge.success{{background:var(--success)}}
.zero-msg{{color:var(--success);font-weight:600;background:#eafaf1;padding:9px 13px;border-radius:8px;font-size:.9em;margin-top:5px}}
.absent-table{{width:100%;border-collapse:collapse;text-align:right;font-size:13px;margin-top:12px}}
.absent-table th{{background:#fdedec;color:var(--danger);padding:10px;font-weight:700}}
.absent-table td{{padding:9px;border-bottom:1px solid #eee;color:#2c3e50}}
.export-btn{{background:var(--success);color:white;border:none;padding:9px 18px;font-family:var(--font);font-size:1em;font-weight:600;border-radius:9px;cursor:pointer;box-shadow:0 3px 8px rgba(0,0,0,0.12);margin-bottom:18px;transition:.2s;display:inline-block}}
.export-btn:hover{{background:#219653;transform:translateY(-2px)}}
.footer{{text-align:center;margin-top:60px;padding:40px;color:#7f8c8d;font-family:'Courier New',monospace;font-size:13px}}

/* ══ MODAL: قائمة الأسماء ══ */
.modal-overlay{{position:fixed;inset:0;background:rgba(0,0,0,0.75);z-index:10000;display:none;align-items:center;justify-content:center;backdrop-filter:blur(4px)}}
.modal-overlay.open{{display:flex}}
.modal{{background:var(--navy2);border:1px solid var(--gold);border-radius:16px;width:min(900px,95vw);height:85vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 24px 80px rgba(0,0,0,0.7);animation:popIn .2s ease}}
.modal-head{{padding:16px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;background:var(--navy3)}}
.modal-head h2{{font-size:15px;font-weight:700;color:var(--gold2)}}
.modal-close{{width:30px;height:30px;border-radius:7px;border:1px solid var(--border);background:transparent;color:var(--text2);font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:.2s}}
.modal-close:hover{{background:rgba(201,168,76,.15);color:var(--gold2)}}
.modal-filters{{padding:10px 16px;border-bottom:1px solid var(--border);display:flex;gap:8px;flex-wrap:wrap;flex-shrink:0}}
.modal-search{{flex:1;min-width:160px;background:var(--navy3);border:1px solid var(--border);border-radius:8px;padding:7px 12px;color:var(--text);font-family:var(--font);font-size:12px;outline:none}}
.modal-search:focus{{border-color:var(--gold)}}
.modal-select{{background:var(--navy3);border:1px solid var(--border);border-radius:8px;padding:7px 10px;color:var(--text);font-family:var(--font);font-size:12px;outline:none;cursor:pointer}}
.modal-count{{padding:8px 16px;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border);flex-shrink:0}}
.modal-count b{{color:var(--gold2)}}
.modal-body{{flex:1;overflow-y:auto;padding:10px 14px}}
.modal-body::-webkit-scrollbar{{width:4px}}
.modal-body::-webkit-scrollbar-thumb{{background:var(--gold);border-radius:2px}}
.modal-table{{width:100%;border-collapse:collapse;font-size:12px}}
.modal-table th{{background:var(--navy3);color:var(--gold2);padding:8px 10px;font-weight:600;text-align:right;position:sticky;top:0;z-index:2;border-bottom:1px solid var(--border)}}
.modal-table td{{padding:7px 10px;color:var(--text);border-bottom:1px solid rgba(255,255,255,0.04)}}
.modal-table tr:hover td{{background:rgba(201,168,76,0.05)}}

/* Stats Modal (الوردية الأولى / الثالثة) */
.stats-modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:50000;align-items:center;justify-content:center;backdrop-filter:blur(5px)}}
.stats-modal-overlay.open{{display:flex}}
.stats-modal-box{{background:var(--navy2);border:1px solid var(--gold);border-radius:16px;width:min(640px,92vw);max-height:85vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 24px 80px rgba(0,0,0,.8);animation:popIn .2s ease}}
.stats-modal-head{{background:linear-gradient(135deg,#0C1F3F,#0F2040);padding:14px 18px;border-bottom:1px solid rgba(201,168,76,.2);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;position:relative}}
.stats-modal-head::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,#C9A84C,transparent)}}
.stats-modal-head h2{{font-size:15px;font-weight:700;color:var(--gold2)}}
.job-card{{background:var(--navy3);border:1px solid rgba(201,168,76,.15);border-radius:9px;margin-bottom:8px;overflow:hidden}}
.job-title{{padding:10px 14px;cursor:pointer;font-weight:600;color:var(--text);display:flex;justify-content:space-between;align-items:center;transition:.15s;font-size:13px}}
.job-title:hover{{background:rgba(201,168,76,.08)}}
.job-badge{{background:#d35400;color:white;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700}}
.emp-list{{display:none;padding:8px 14px 10px;border-top:1px solid rgba(255,255,255,.06);list-style:disc;padding-right:30px;background:var(--navy)}}
.emp-list li{{padding:3px 0;color:var(--text2);font-size:12px;border-bottom:1px dashed rgba(255,255,255,.05)}}
.emp-list li:last-child{{border:none}}

/* ── FOOTER BAR ── */
.ftr{{flex-shrink:0;background:var(--navy2);border-top:1px solid var(--border);padding:6px 24px;display:flex;align-items:center;justify-content:space-between;font-size:10.5px;color:var(--text2)}}
.ftr-pulse{{width:6px;height:6px;border-radius:50%;background:var(--teal2);animation:pulse-anim 2s infinite}}
</style>
</head>
<body>

<!-- ══ HEADER ══ -->
<div class="hdr">
  <div class="hdr-brand">
    <div class="hdr-logo">🕋</div>
    <div>
      <div class="hdr-title">لوحة التحكم الميدانية – حج 1447</div>
      <div class="hdr-sub">المسجد الحرام والساحات المحيطة – شركة العربة الفاخرة</div>
    </div>
  </div>
  <div class="hdr-center">
    <div class="kpi"><span class="kpi-num">{total_employees}</span><span class="kpi-lbl">إجمالي الموظفين</span></div>
    <div class="kpi"><span class="kpi-num">{permanent_count}</span><span class="kpi-lbl">دائمين</span></div>
    <div class="kpi"><span class="kpi-num">{seasonal_count}</span><span class="kpi-lbl">موسميين</span></div>
    <div class="kpi"><span class="kpi-num">{total_companies}</span><span class="kpi-lbl">شركات</span></div>
    <div class="kpi"><span class="kpi-num">{total_map}</span><span class="kpi-lbl">في الخريطة</span></div>
  </div>
  <div class="hdr-right">
    <div class="hdr-time"><b>{now_str}</b>توقيت مكة</div>
    <button class="icon-btn" onclick="openStatsModal('الوردية الأولى')" title="إحصاء الوردية الأولى">☀️</button>
    <button class="icon-btn" onclick="openStatsModal('الوردية الثالثة')" title="إحصاء الوردية الثالثة">🌙</button>
    <button class="icon-btn" onclick="openNamesModal()" title="قائمة جميع الأسماء">👥</button>
    <button class="icon-btn" onclick="toggleFS()" title="ملء الشاشة">⛶</button>
  </div>
</div>

<!-- ══ NAV TABS ══ -->
<div class="nav-tabs">
  <button class="tab-btn active" onclick="switchTab('map',this)">🗺️ الخريطة الميدانية</button>
  <button class="tab-btn" onclick="switchTab('staff',this)">🏢 القوى العاملة</button>
  <button class="tab-btn" onclick="switchTab('absent',this)">🚨 سجل الغياب ({active_shift_lbl})</button>
</div>

<!-- ══ PAGES ══ -->

<!-- TAB 1: الخريطة الحقيقية -->
<div class="tab-page active" id="page-map">
  <div class="map-panel" style="display:flex;flex-direction:column">
    <!-- sub-tabs للورديات داخل تبويب الخريطة -->
    <div class="shift-subtabs">
      <button class="stab active" onclick="setShift('all',this)">🗺️ الكل</button>
      <button class="stab" onclick="setShift('s1',this)">☀️ الأولى</button>
      <button class="stab" onclick="setShift('s3',this)">🌙 الثالثة</button>
    </div>
    <div class="panel-label">ملخص الكوادر</div>
    <div class="mini-stats">
      <div class="mini-stat"><div class="n">{total_map}</div><div class="l">إجمالي الخريطة</div></div>
      <div class="mini-stat"><div class="n">{total_sites}</div><div class="l">موقع</div></div>
      <div class="mini-stat"><div class="n">{total_s1}</div><div class="l">وردية أولى</div></div>
      <div class="mini-stat"><div class="n">{total_s3}</div><div class="l">وردية ثالثة</div></div>
    </div>
    <div class="panel-label-wrap">
      <span class="panel-label" style="padding:0">المواقع الميدانية</span>
      <span class="site-count-badge" id="siteBadge">0 موقع</span>
    </div>
    <div class="search-wrap">
      <input class="search-input" id="searchBox" placeholder="ابحث عن موقع..." oninput="filterList(this.value)">
    </div>
    <div class="site-list" id="siteList"></div>
  </div>
  <div class="map-wrap"><div id="map"></div></div>
</div>

<!-- TAB 2: القوى العاملة -->
<div class="tab-page-scroll" id="page-staff">
  <div class="report-page">
    <div class="rp-header">
      <div class="eng-badge" dir="ltr"><span class="title">Eng.</span><span class="name">Abdulaziz Alshehri</span></div>
      <h1>التقرير الشامل لموسم حج 1447</h1>
      <div style="font-size:1.1em;margin-top:8px;color:#e0f2f1">متابعة ميدانية حية</div>
      <div class="live-indicator"><span class="pulse"></span>آخر تحديث: {now_str}</div>
    </div>
    <div class="stats-container">
      <div class="stat-card"><b>{total_employees}</b><span>إجمالي الفعالين</span></div>
      <div class="stat-card"><b>{permanent_count}</b><span>موظفين دائمين</span></div>
      <div class="stat-card"><b>{seasonal_count}</b><span>موظفين موسميين</span></div>
      <div class="stat-card"><b>{total_companies}</b><span>الشركات المشغلة</span></div>
      <div class="stat-card"><b>{total_shifts}</b><span>الورديات المجدولة</span></div>
    </div>
    <h2 class="section-title">🏢 القوى العاملة (جميع الموظفين لكل الشركات)</h2>
    {companies_html}
    <h2 class="section-title">⏱ إجمالي كل وردية (جميع الشركات)</h2>
    <div class="shifts-totals-grid">{shifts_totals_html}</div>
    <div class="grand-summary">
      <h2>📊 الإجمالي الشامل للوظائف (كافة الشركات)</h2>
      <ul class="grand-jobs-list">{grand_jobs_html}</ul>
    </div>
    <div class="footer" dir="ltr">AUTOMATED SYSTEM PREPARED BY<br>
      <b style="color:var(--primary);font-size:1.4em">Eng. Abdulaziz Alshehri</b></div>
  </div>
</div>

<!-- TAB 3: الغياب -->
<div class="tab-page-scroll" id="page-absent">
  <div class="report-page">
    <div class="grand-summary" style="border-color:var(--danger);margin-top:0;margin-bottom:30px">
      <h2 style="color:var(--danger)">🚨 سجل الغياب – {active_shift_lbl}</h2>
      <div style="text-align:center;margin-bottom:20px">
        <button class="export-btn" onclick="downloadCSV()">📥 تحميل كشف الغياب (CSV)</button>
      </div>
      {absent_html}
    </div>
    <div class="footer" dir="ltr">AUTOMATED SYSTEM PREPARED BY<br>
      <b style="color:var(--primary);font-size:1.4em">Eng. Abdulaziz Alshehri</b></div>
  </div>
</div>

<!-- ══ MODAL: قائمة الأسماء الكاملة ══ -->
<div class="modal-overlay" id="modalOverlay" onclick="closeModalOutside(event)">
  <div class="modal">
    <div class="modal-head">
      <h2>👥 قائمة جميع الموظفين ({total_employees} موظف)</h2>
      <button class="modal-close" onclick="closeNamesModal()">✕</button>
    </div>
    <div class="modal-filters">
      <input class="modal-search" id="modalSearch" placeholder="ابحث بالاسم أو الهوية..." oninput="filterNamesModal()">
      <select class="modal-select" id="filterCompany" onchange="filterNamesModal()"><option value="">كل الشركات</option></select>
      <select class="modal-select" id="filterShift" onchange="filterNamesModal()"><option value="">كل الورديات</option></select>
      <select class="modal-select" id="filterJob" onchange="filterNamesModal()"><option value="">كل الوظائف</option></select>
    </div>
    <div class="modal-count" id="modalCount"></div>
    <div class="modal-body">
      <table class="modal-table">
        <thead><tr><th>#</th><th>الاسم</th><th>الهوية</th><th>الوظيفة</th><th>الشركة</th><th>الوردية</th></tr></thead>
        <tbody id="modalBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ══ MODAL: إحصاء الورديات ══ -->
<div class="stats-modal-overlay" id="statsModal" onclick="closeStatsOnBg(event)">
  <div class="stats-modal-box">
    <div class="stats-modal-head">
      <h2 id="statsModalTitle">إحصاء الوردية</h2>
      <button class="modal-close" onclick="closeStatsModal()">✕</button>
    </div>
    <div class="modal-filters">
      <input class="modal-search" id="statsSearch" placeholder="ابحث عن وظيفة أو اسم..." oninput="filterStatsModal(this.value)">
    </div>
    <div class="modal-body" id="statsModalBody"></div>
  </div>
</div>

<!-- ══ FOOTER ══ -->
<div class="ftr">
  <div style="display:flex;align-items:center;gap:7px">
    <span class="ftr-pulse"></span>
    <span>☀️ الأولى (8ص–4م) &nbsp;|&nbsp; 🌙 الثالثة (4م–12م) &nbsp;|&nbsp; اضغط على الدبابيس لعرض التفاصيل &nbsp;|&nbsp; 👥 لقائمة الأسماء</span>
  </div>
  <span>تطوير م. عبدالعزيز الشهري &nbsp;|&nbsp; شركة العربة الفاخرة</span>
</div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
/* ═══════════════════════════════════════
   TAB SWITCHING
═══════════════════════════════════════ */
function switchTab(name, btn) {{
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-page,.tab-page-scroll').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('page-' + name).classList.add('active');
  if (name === 'map') setTimeout(() => map.invalidateSize(), 50);
}}

function toggleFS() {{
  if (!document.fullscreenElement) document.documentElement.requestFullscreen().catch(()=>{{}});
  else document.exitFullscreen();
}}
document.addEventListener('keydown', e => {{ if(e.key==='F11'){{e.preventDefault();toggleFS();}} }});

/* ═══════════════════════════════════════
   MAP (الخريطة الحقيقية من tanaqol_map_final)
═══════════════════════════════════════ */
const PINS       = {pins_json};
const STATS_DATA = {stats_json};

const map = L.map('map', {{center:[21.4225,39.8262], zoom:17, zoomControl:false}});
L.tileLayer('https://mt1.google.com/vt/lyrs=y&x={{x}}&y={{y}}&z={{z}}',
  {{attribution:'© Google Maps', maxZoom:20}}).addTo(map);
L.control.zoom({{position:'bottomleft'}}).addTo(map);

function buildPopup(p) {{
  const rows = p.employees.map((e,i) => `
    <tr>
      <td style="color:#C9A84C;font-weight:600">${{i+1}}</td>
      <td>${{e.source==='staff_data'?'<span class="source-tag">staff</span>':''}}${{e.name}}</td>
      <td style="color:#555">${{e.job}}</td>
      <td dir="ltr" style="text-align:right;color:#777">${{e.phone}}</td>
    </tr>`).join('');
  return `<div class="pop">
    <div class="pop-head">
      <div class="pop-title">📍 ${{p.site}}</div>
      <div class="pop-shift">${{p.shift}}</div>
    </div>
    <div class="pop-mgmt">
      <div class="pop-row"><b>💼 رئيس الأخصائيين:</b><span class="val">${{p.chief_hr}}</span></div>
      <div class="pop-row"><b>🎖️ رئيس الوردية:</b><span class="val">${{p.field_ldr}}</span></div>
      <div class="pop-row"><b>👑 مساعد المشرف:</b><span class="val">${{p.asst_ldr}}</span></div>
      <div class="pop-row"><b>🌟 أخصائي الموارد:</b><span class="val gold">${{p.hr_spec}}</span></div>
      <div class="pop-row"><b>📌 المشرفين:</b><span class="val teal">${{p.supervisors}}</span></div>
    </div>
    <div class="pop-emps">
      <div class="pop-emps-title">طاقم الموقع (${{p.total}} موظف)</div>
      <div class="pop-scroll">
        <table class="pop-table">
          <thead><tr><th>#</th><th>الاسم</th><th>الوظيفة</th><th>الجوال</th></tr></thead>
          <tbody>${{rows}}</tbody>
        </table>
      </div>
    </div>
  </div>`;
}}

const markers   = {{}};
const sideItems = {{}};
let activeKey   = null;
let activeShift = 'all';

function pinKey(p) {{ return p.shift + '|' + p.site; }}

function setActive(key) {{
  if (activeKey) {{
    if (sideItems[activeKey]) sideItems[activeKey].classList.remove('active');
    const old = document.querySelector(`.pin-body[data-key="${{activeKey}}"]`);
    if (old) old.classList.remove('active');
  }}
  activeKey = key;
  if (sideItems[key]) sideItems[key].classList.add('active');
  const pin = document.querySelector(`.pin-body[data-key="${{key}}"]`);
  if (pin) pin.classList.add('active');
}}

PINS.forEach(p => {{
  const key  = pinKey(p);
  const isS1 = p.shift === 'الوردية الأولى';
  const cls  = isS1 ? 'pin-s1' : 'pin-s3';
  const tCls = isS1 ? 'pin-s1-tail' : 'pin-s3-tail';

  const icon = L.divIcon({{
    html: `<div class="pin-wrap">
      <div class="pin-body ${{cls}}" data-key="${{key}}">${{p.total}}</div>
      <div class="pin-tail ${{tCls}}"></div>
    </div>`,
    className:'', iconSize:[40,52], iconAnchor:[20,52], popupAnchor:[0,-54]
  }});

  const mk = L.marker([p.lat, p.lon], {{icon}})
    .addTo(map)
    .bindPopup(buildPopup(p), {{maxWidth:400, className:'', closeButton:true}});
  mk.on('click', () => setActive(key));
  markers[key] = {{marker:mk, shift:p.shift}};

  const item = document.createElement('div');
  item.className = 'sitem';
  item.dataset.name  = p.site;
  item.dataset.shift = p.shift;
  item.innerHTML = `
    <div class="sitem-top">
      <div class="sitem-name">${{p.site}}</div>
      <div class="epill">${{p.total}}</div>
    </div>
    <div class="sitem-shift">${{isS1?'☀️':'🌙'}} ${{p.shift}}</div>`;
  item.onclick = () => {{
    setActive(key);
    map.flyTo([p.lat, p.lon], 18, {{animate:true, duration:.85}});
    setTimeout(() => mk.openPopup(), 900);
    item.scrollIntoView({{behavior:'smooth', block:'nearest'}});
  }};
  document.getElementById('siteList').appendChild(item);
  sideItems[key] = item;
}});
updateSideCount();

function setShift(sh, btn) {{
  document.querySelectorAll('.stab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  activeShift = sh;
  Object.entries(markers).forEach(([key, {{marker, shift}}]) => {{
    const show = sh==='all' || (sh==='s1'&&shift==='الوردية الأولى') || (sh==='s3'&&shift==='الوردية الثالثة');
    if (show) map.addLayer(marker); else map.removeLayer(marker);
  }});
  document.querySelectorAll('.sitem').forEach(el => {{
    const match = sh==='all' || (sh==='s1'&&el.dataset.shift==='الوردية الأولى') || (sh==='s3'&&el.dataset.shift==='الوردية الثالثة');
    el.style.display = match ? '' : 'none';
  }});
  updateSideCount();
}}

function filterList(q) {{
  const t = q.trim();
  document.querySelectorAll('.sitem').forEach(el => {{
    const shiftOk = activeShift==='all' || (activeShift==='s1'&&el.dataset.shift==='الوردية الأولى') || (activeShift==='s3'&&el.dataset.shift==='الوردية الثالثة');
    const nameOk  = !t || el.dataset.name.includes(t);
    el.style.display = (shiftOk && nameOk) ? '' : 'none';
  }});
  updateSideCount();
}}

function updateSideCount() {{
  const n = [...document.querySelectorAll('.sitem')].filter(e=>e.style.display!=='none').length;
  document.getElementById('siteBadge').textContent = n + ' موقع';
}}

/* ═══════════════════════════════════════
   Stats Modal (إحصاء الورديات)
═══════════════════════════════════════ */
let _currentStatsData = null;

function openStatsModal(shiftName) {{
  const data = STATS_DATA[shiftName];
  if (!data) return;
  _currentStatsData = data;
  document.getElementById('statsModalTitle').textContent =
    (shiftName==='الوردية الأولى'?'☀️ ':'🌙 ') + shiftName + '  —  إجمالي: ' + data.total + ' موظف';
  renderStatsModal(data.jobs);
  document.getElementById('statsSearch').value = '';
  document.getElementById('statsModal').classList.add('open');
}}

function renderStatsModal(jobs) {{
  let html = '';
  for (const job in jobs) {{
    const emps = jobs[job];
    const lis  = emps.map(e=>`<li>${{e}}</li>`).join('');
    html += `<div class="job-card">
      <div class="job-title" onclick="toggleJobList(this)">
        <span>🛠️ ${{job}}</span>
        <span class="job-badge">${{emps.length}} موظف ▾</span>
      </div>
      <ul class="emp-list">${{lis}}</ul>
    </div>`;
  }}
  document.getElementById('statsModalBody').innerHTML =
    html || '<p style="color:var(--text2);text-align:center;padding:20px">لا توجد بيانات</p>';
}}

function filterStatsModal(q) {{
  if (!_currentStatsData) return;
  const t = q.trim();
  if (!t) {{ renderStatsModal(_currentStatsData.jobs); return; }}
  const filtered = {{}};
  for (const job in _currentStatsData.jobs) {{
    if (job.includes(t)) {{ filtered[job] = _currentStatsData.jobs[job]; continue; }}
    const names = _currentStatsData.jobs[job].filter(n => n.includes(t));
    if (names.length) filtered[job] = names;
  }}
  renderStatsModal(filtered);
}}

function toggleJobList(el) {{
  const ul    = el.nextElementSibling;
  const badge = el.querySelector('.job-badge');
  const open  = ul.style.display === 'block';
  ul.style.display = open ? 'none' : 'block';
  badge.textContent = badge.textContent.replace(open?'▲':'▾', open?'▾':'▲');
}}

function closeStatsModal() {{ document.getElementById('statsModal').classList.remove('open'); }}
function closeStatsOnBg(e) {{ if (e.target===document.getElementById('statsModal')) closeStatsModal(); }}

/* ═══════════════════════════════════════
   Names Modal (قائمة الأسماء الكاملة)
═══════════════════════════════════════ */
const ALL_EMPS = {all_names_json};

const _companies=[...new Set(ALL_EMPS.map(e=>e.company).filter(Boolean))].sort();
const _shifts   =[...new Set(ALL_EMPS.map(e=>e.shift).filter(Boolean))].sort();
const _jobs     =[...new Set(ALL_EMPS.map(e=>e.job).filter(Boolean))].sort();
_companies.forEach(c=>{{document.getElementById('filterCompany').innerHTML+=`<option value="${{c}}">${{c}}</option>`}});
_shifts.forEach(s=>{{document.getElementById('filterShift').innerHTML+=`<option value="${{s}}">${{s}}</option>`}});
_jobs.forEach(j=>{{document.getElementById('filterJob').innerHTML+=`<option value="${{j}}">${{j}}</option>`}});

function filterNamesModal() {{
  const q   =document.getElementById('modalSearch').value.trim().toLowerCase();
  const comp=document.getElementById('filterCompany').value;
  const sh  =document.getElementById('filterShift').value;
  const jo  =document.getElementById('filterJob').value;
  const filtered=ALL_EMPS.filter(e=>
    (!q||(e.name&&e.name.toLowerCase().includes(q))||(e.id&&e.id.includes(q)))&&
    (!comp||e.company===comp)&&(!sh||e.shift===sh)&&(!jo||e.job===jo)
  );
  document.getElementById('modalCount').innerHTML=`يُعرض <b>${{filtered.length}}</b> من أصل <b>${{ALL_EMPS.length}}</b> موظف`;
  document.getElementById('modalBody').innerHTML=filtered.map((e,i)=>
    `<tr><td style="color:var(--gold2);font-weight:600">${{i+1}}</td><td>${{e.name||''}}</td><td dir="ltr" style="text-align:right;color:var(--text2)">${{e.id||''}}</td><td>${{e.job||''}}</td><td>${{e.company||''}}</td><td>${{e.shift||''}}</td></tr>`
  ).join('');
}}

function openNamesModal() {{
  document.getElementById('modalOverlay').classList.add('open');
  filterNamesModal();
}}
function closeNamesModal() {{ document.getElementById('modalOverlay').classList.remove('open'); }}
function closeModalOutside(e) {{ if(e.target===document.getElementById('modalOverlay')) closeNamesModal(); }}
document.addEventListener('keydown', e=>{{ if(e.key==='Escape'){{ closeNamesModal(); closeStatsModal(); }} }});

/* ═══════════════════════════════════════
   EXPORT CSV
═══════════════════════════════════════ */
function downloadCSV() {{
  var csv=`{csv_data}`;
  var blob=new Blob(["\\uFEFF"+csv],{{type:'text/csv;charset=utf-8;'}});
  var url=URL.createObjectURL(blob);
  var a=document.createElement("a");
  a.href=url; a.download="سجل_الغياب_{today_iso}.csv";
  a.style.visibility='hidden';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}}
</script>
</body>
</html>"""

    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(HTML)

    print("=" * 56)
    print("✅  تم إنشاء اللوحة الموحدة بنجاح!")
    print(f"📁  الملف         : {REPORT_PATH}")
    print(f"🗺️   في الخريطة   : {total_map}  |  الأولى: {total_s1}  |  الثالثة: {total_s3}")
    print(f"👥  القوى العاملة : {total_employees}")
    print(f"🏢  الشركات       : {total_companies}")
    print(f"⏱   الورديات      : {total_shifts}")
    print(f"📍  المواقع       : {total_sites}")
    print(f"➕  مُضاف من staff : {len(staff_unregistered)}")
    print("=" * 56)
    webbrowser.open("file:///" + REPORT_PATH.replace("\\", "/"))

if __name__ == "__main__":
    generate()
