import time
import os
import requests
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# --- الإعدادات الأساسية ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_PATH = os.path.join(BASE_DIR, "index.html")
EXCEL_PATH = os.path.join(BASE_DIR, "staff_data.xlsx") 

USERNAME = os.getenv('HAJJ_USER')
PASSWORD = os.getenv('HAJJ_PASS')

def safe_extract_list(res_json):
    if not res_json: return []
    if isinstance(res_json, list): return res_json
    if isinstance(res_json, dict):
        data = res_json.get('data')
        if isinstance(data, list): return data
        if isinstance(data, dict): return data.get('list', [])
    return []

def get_hajj_token():
    driver = None
    try:
        print("🌐 [1/4] جاري فتح المتصفح الخفي لسحب التوكن...")
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        driver.get("https://tnql-prod.sejeltech.app/")
        time.sleep(5) 
        if "404" in driver.current_url:
            try:
                driver.find_element(By.XPATH, "//*[contains(text(), 'HOME')]").click()
                time.sleep(5)
            except: pass
        
        wait = WebDriverWait(driver, 15)
        user_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input)[1]")))
        pass_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
        
        user_input.send_keys(USERNAME)
        pass_input.send_keys(PASSWORD)
        pass_input.send_keys(Keys.RETURN)
        time.sleep(12) 
        
        token = driver.execute_script("return window.localStorage.getItem('token')") or \
                next((v for k, v in driver.execute_script("return window.localStorage;").items() if str(v).startswith("eyJ")), None)
        
        if token:
            print("✅ [4/4] تم سحب التوكن الآلي بنجاح!")
            return f"Bearer {token}"
        return None
    except Exception as e:
        print(f"❌ خطأ المتصفح: {e}")
        return None
    finally:
        if driver: driver.quit()

def generate_master_dashboard():
    # 🎯 توقيت مكة المكرمة
    makkah_time = datetime.utcnow() + timedelta(hours=3)
    current_time_str = makkah_time.strftime("%Y-%m-%d %I:%M %p")
    hour = makkah_time.hour
    
    # 🎯 الكلمات المفتاحية الذكية بناءً على المسميات الدقيقة
    if 8 <= hour < 16:
        shift_keywords = ["ثاني", "صباح", "2"]
        active_shift_title = "الوردية الثانية صباح"
    elif 16 <= hour < 24:
        shift_keywords = ["ثالث", "مساء", "3"]
        active_shift_title = "وردية الثالثة مساء"
    else:
        shift_keywords = ["اول", "ليل", "1"]
        active_shift_title = "الوردية الاولى ليلية"

    # 1. تجهيز قاعدة بيانات الإكسيل
    excel_db = {}
    if os.path.exists(EXCEL_PATH):
        try:
            df_excel = pd.read_excel(EXCEL_PATH, dtype=str).fillna('غير متوفر')
            
            phone_col = 'غير متوفر'
            for col in ['رقم التليفون', 'رقم الجوال', 'الجوال', 'Phone', 'Mobile']:
                if col in df_excel.columns:
                    phone_col = col
                    break

            for _, row in df_excel.iterrows():
                n_id = str(row.get('هويه الموظف', '')).replace('.0', '').strip().lower()
                if n_id and n_id != 'غير متوفر' and n_id != 'nan':
                    excel_db[n_id] = {
                        'name': str(row.get('اسم الموظف', 'غير متوفر')).strip(), 
                        'job': str(row.get('الوظيفه', 'غير متوفر')).strip(),
                        'dept': str(row.get('القسم', 'غير متوفر')).strip(),
                        'company': str(row.get('شركة التشغيل', 'غير متوفر')).strip(),
                        'phone': str(row.get(phone_col, 'غير متوفر')).strip() if phone_col != 'غير متوفر' else 'غير متوفر'
                    }
        except Exception as e:
            pass

    token = get_hajj_token()
    if not token: 
        return

    headers = {"authorization": token, "content-type": "application/json", "lang": "ar"}
    payload = {
        "paging": {"sortField": "Id", "searchOrder": 2, "pageIndex": 1, "totalRowsCount": 10469, "totalPages": 1, "pageSize": 11000, "sortBy": "Id Desc"},
        "data": {
            "searchText": "", "name": "", "EmployeeId": None, "OccupationIds": [], "DepartmentIds": [], 
            "SectionIds": [], "WorkShiftIds": [], "EmployeeTypes": [], "ManagerIds": [], 
            "OperatorCompanyIds": [], "NationalIdExpired": [], "ActiveStatus": [True], 
            "isPrinted": None, "isDeleted": False
        }
    }

    try:
        r_emp = requests.post("https://tnql-prod.sejeltech.app/api/StaffMember/GetStaffMember", headers=headers, json=payload)
        all_employees = safe_extract_list(r_emp.json())
        
        r_att = requests.post("https://tnql-prod.sejeltech.app/api/EmployeeAttendanceMonitor/GetAttendance", headers=headers, json=payload)
        att_data = safe_extract_list(r_att.json())
        
        # 🔥 فلترة الحضور لتاريخ اليوم فقط
        today_iso = makkah_time.strftime("%Y-%m-%d") 
        today_ar1 = makkah_time.strftime("%d/%m/%Y") 
        today_ar2 = f"{makkah_time.day}/{makkah_time.month}/{makkah_time.year}" 

        present_ids = set()
        for x in att_data:
            if isinstance(x, dict):
                row_str = str(x.values()).lower()
                if today_iso in row_str or today_ar1 in row_str or today_ar2 in row_str:
                    if "غائب" not in row_str and "absent" not in row_str:
                        emp_code = x.get('employeeCode')
                        if emp_code:
                            present_ids.add(str(emp_code).strip().lower())

        if not all_employees:
            return

        for emp in all_employees:
            nid = str(emp.get('nationalId', '')).replace('.0', '').strip().lower()
            api_phone = emp.get('mobileNumber') or emp.get('phoneNumber') or 'لا يوجد رقم'
            
            if nid in excel_db:
                ex = excel_db[nid]
                if ex['job'] not in ['غير متوفر', 'nan']: emp['occupationName'] = ex['job']
                if ex['company'] not in ['غير متوفر', 'nan']: emp['operatorCompanyName'] = ex['company']
                emp['clean_name'] = ex['name'] if ex['name'] not in ['غير متوفر', 'nan'] else (emp.get('name') or 'غير متوفر')
                emp['clean_dept'] = ex['dept']
                emp['clean_phone'] = ex['phone'] if ex['phone'] not in ['غير متوفر', 'nan'] else api_phone
            else:
                emp['clean_name'] = emp.get('name') or 'غير متوفر'
                emp['clean_dept'] = 'غير متوفر'
                emp['clean_phone'] = api_phone

        df = pd.DataFrame(all_employees)
        df = df.fillna('غير محدد')

        total_employees = len(df)
        total_companies = df['operatorCompanyName'].nunique()
        total_shifts = df['workShiftName'].nunique()
        
        def clean_type(val):
            v = str(val).lower()
            return 'دائم' if 'permanent' in v or 'دائم' in v else ('موسمي' if 'seasonal' in v or 'موسمي' in v else 'غير محدد')
        
        df['mapped_type'] = df.get('employeeTypeName', '').apply(clean_type)
        permanent_count = len(df[df['mapped_type'] == 'دائم'])
        seasonal_count = len(df[df['mapped_type'] == 'موسمي'])

        # --- بداية التصميم (HTML/CSS) ---
        html_content = f"""<!DOCTYPE html>
        <html lang="ar" dir="rtl">
        <head>
            <meta charset="UTF-8">
            <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap" rel="stylesheet">
            <style>
                :root {{ --primary: #004d40; --secondary: #00796b; --accent: #c0a16b; --bg: #f4f7f6; --danger: #e74c3c; --success: #27ae60; }}
                body {{ font-family: 'Cairo', sans-serif; background-color: var(--bg); margin: 0; padding: 20px; color: #2c3e50; }}
                .header {{ text-align: center; background: linear-gradient(135deg, var(--primary), var(--secondary)); padding: 60px 20px; border-radius: 30px; color: white; box-shadow: 0 15px 35px rgba(0,0,0,0.2); position: relative; }}
                .eng-badge {{ display: inline-block; border: 2px solid var(--accent); padding: 12px 35px; border-radius: 15px; margin-bottom: 25px; background: rgba(0,0,0,0.3); }}
                .eng-badge .title {{ display: block; font-size: 1.2em; color: var(--accent); font-weight: 900; letter-spacing: 2px; }}
                .eng-badge .name {{ display: block; font-size: 1.8em; font-weight: 700; }}
                h1 {{ font-size: 2.5em; margin: 10px 0; }}
                .live-indicator {{ background: rgba(255,255,255,0.15); padding: 8px 20px; border-radius: 50px; font-weight: 700; display: inline-flex; align-items: center; gap: 10px; margin-top: 15px; }}
                .pulse {{ height: 12px; width: 12px; background-color: #4caf50; border-radius: 50%; animation: pulse-animation 1.5s infinite; }}
                @keyframes pulse-animation {{ 0% {{ box-shadow: 0 0 0 0px rgba(76, 175, 80, 0.7); }} 100% {{ box-shadow: 0 0 0 10px rgba(76, 175, 80, 0); }} }}
                .stats-container {{ display: flex; gap: 15px; justify-content: center; margin: -40px 0 50px; flex-wrap: wrap; position: relative; }}
                .stat-card {{ background: white; padding: 20px 15px; border-radius: 20px; text-align: center; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border-bottom: 6px solid var(--accent); flex: 1; min-width: 150px; max-width: 200px; }}
                .stat-card b {{ display: block; font-size: 2.5em; color: var(--primary); margin-bottom: 5px; }}
                .stat-card span {{ font-size: 0.95em; font-weight: 700; color: #7f8c8d; }}
                .company-card {{ background: white; padding: 35px; border-radius: 30px; margin-bottom: 40px; box-shadow: 0 15px 40px rgba(0,0,0,0.05); border: 1px solid #eee; }}
                .company-title {{ font-size: 2em; color: var(--primary); border-bottom: 3px solid #f0f0f0; padding-bottom: 20px; margin-bottom: 25px; font-weight: 900; }}
                .shift-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 25px; }}
                .shift-box {{ background: #fafafa; border: 1px solid #eef2f3; padding: 25px; border-radius: 20px; }}
                .shift-name {{ color: var(--secondary); font-weight: 800; font-size: 1.3em; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; border-right: 5px solid var(--accent); padding-right: 15px; }}
                .shift-total-badge {{ font-size: 0.75em; background: #e0f2f1; color: var(--primary); padding: 5px 12px; border-radius: 12px; font-weight: 900; }}
                .jobs-list {{ list-style: none; padding: 0; margin: 0; }}
                .jobs-list li {{ display: block; padding: 10px 0; border-bottom: 1px dashed #ddd; font-weight: 700; }}
                .job-summary {{ display: flex; justify-content: space-between; align-items: center; cursor: pointer; list-style: none; outline: none; }}
                .job-summary::-webkit-details-marker {{ display: none; }}
                .job-val {{ background: var(--primary); color: white; padding: 3px 12px; border-radius: 8px; font-family: monospace; font-size: 1.1em; }}
                .emp-details-list {{ margin-top: 12px; padding: 15px; background: rgba(0, 77, 64, 0.04); border-radius: 10px; border-right: 4px solid var(--accent); font-size: 0.85em; list-style: none; }}
                .emp-details-list li {{ display: flex; flex-direction: column; gap: 5px; padding: 8px 0; border-bottom: 1px solid rgba(0,0,0,0.05); color: #2c3e50; font-weight: 600; border-bottom-style: solid; }}
                .emp-details-list li:last-child {{ border-bottom: none; padding-bottom: 0; }}
                .emp-name {{ color: var(--primary); font-weight: 800; font-size: 1.1em; }}
                .emp-meta {{ color: #7f8c8d; display: flex; justify-content: space-between; font-family: Tahoma, sans-serif; }}
                .grand-summary {{ background: #ffffff; border: 4px solid var(--primary); padding: 40px; border-radius: 40px; margin-top: 60px; }}
                .grand-summary h2 {{ text-align: center; color: var(--primary); font-size: 2.2em; margin-bottom: 35px; font-weight: 900; }}
                .grand-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; }}
                .grand-item {{ background: #e0f2f1; padding: 15px; border-radius: 12px; display: flex; justify-content: space-between; align-items: center; font-weight: 800; border-right: 5px solid var(--primary); }}
                .absent-table {{ width: 100%; border-collapse: collapse; text-align: right; margin-top: 15px; font-size: 14px; }}
                .absent-table th {{ background: #f9ebec; color: var(--danger); padding: 12px; }}
                .absent-table td {{ padding: 10px; border-bottom: 1px solid #eee; }}
                .footer {{ text-align: center; margin-top: 80px; padding: 50px; color: #7f8c8d; font-family: 'Courier New', monospace; }}
                .zero-absent-msg {{ color: var(--success); font-weight: bold; text-align: right; margin-top: 5px; background: #e8f8f5; padding: 10px; border-radius: 8px; font-size: 0.9em; }}
                .export-btn {{ background: #27ae60; color: white; border: none; padding: 10px 20px; font-family: 'Cairo', sans-serif; font-size: 1.1em; font-weight: bold; border-radius: 10px; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; display: inline-block; transition: 0.2s; }}
                .export-btn:hover {{ background: #219653; transform: translateY(-2px); }}
            </style>
        </head>
        <body>
            <div class="header">
                <div class="eng-badge" dir="ltr"><span class="title">Eng.</span><span class="name">Abdulaziz Alshehri</span></div>
                <h1>التقرير الشامل لموسم حج 1447</h1>
                <div style="font-size: 1.2em; margin-top: 10px; color: #e0f2f1;">متابعة ميدانية حية لـ: {active_shift_title}</div>
                <div class="live-indicator"><span class="pulse"></span>آخر تحديث تلقائي: {current_time_str}</div>
            </div>

            <div class="stats-container">
                <div class="stat-card"><b>{total_employees}</b><span>إجمالي الفعالين</span></div>
                <div class="stat-card"><b>{permanent_count}</b><span>موظفين دائمين</span></div>
                <div class="stat-card"><b>{seasonal_count}</b><span>موظفين موسميين</span></div>
                <div class="stat-card"><b>{total_companies}</b><span>الشركات المشغلة</span></div>
                <div class="stat-card"><b>{total_shifts}</b><span>الورديات المجدولة</span></div>
            </div>

            <div class="content">
                <h2 style='text-align:center; color:var(--primary); margin: 50px 0 30px; font-size: 2.2em;'>🏢 القوى العاملة (جميع الموظفين لكل الشركات)</h2>
        """

        # --- 4. القسم العلوي: جميع الموظفين لكل الشركات (بالتفصيل القديم اللي يعجبك) ---
        for company in df['operatorCompanyName'].unique():
            if pd.isna(company) or company == 'غير محدد': continue
            html_content += f"<div class='company-card'><div class='company-title'>🏢 {company}</div><div class='shift-grid'>"
            
            for shift in df[df['operatorCompanyName']==company]['workShiftName'].unique():
                shift_df = df[(df['operatorCompanyName']==company) & (df['workShiftName']==shift)]
                shift_total = len(shift_df)
                shift_jobs = shift_df['occupationName'].value_counts()
                
                jobs_html = ""
                for job_name, count in shift_jobs.items():
                    job_df = shift_df[shift_df['occupationName'] == job_name]
                    emp_list_html = ""
                    for _, row in job_df.iterrows():
                        e_name = row.get('clean_name', 'غير متوفر')
                        e_id = str(row.get('nationalId', '')).replace('.0', '')
                        e_phone = row.get('clean_phone', 'لا يوجد رقم')
                        
                        emp_list_html += f"<li><span class='emp-name'>👤 {e_name}</span><span class='emp-meta'><span>💳 {e_id}</span> <span>📱 {e_phone}</span></span></li>"
                    
                    jobs_html += f"<li><details><summary class='job-summary'><span>{job_name}</span><span class='job-val'>{count} ▾</span></summary><ul class='emp-details-list'>{emp_list_html}</ul></details></li>"
                
                html_content += f"<div class='shift-box'><span class='shift-name'><span>📍 {shift}</span><span class='shift-total-badge'>العدد: {shift_total}</span></span><ul class='jobs-list'>{jobs_html}</ul></div>"
            html_content += "</div></div>"

        # --- 4.5 الإضافة الجديدة: الإجمالي العام للوظائف (لكافة الشركات) ---
        total_job_counts = df['occupationName'].value_counts()
        jobs_summary_html = "<div class='grand-summary' style='padding: 30px; margin-bottom: 40px;'><h2 style='text-align:center; color:var(--primary); font-size: 2em; margin-bottom: 25px;'>📊 الإجمالي الشامل للوظائف (كافة الشركات)</h2><div class='grand-grid'>"
        for j, v in total_job_counts.items():
            jobs_summary_html += f"<div class='grand-item'><span style='color:var(--secondary); font-size:1.1em;'>{j}</span><span class='job-val' style='font-size:1.3em;'>{v}</span></div>"
        jobs_summary_html += "</div></div>"
        
        html_content += jobs_summary_html

        # --- 5. قسم الغياب: ذكي وموثوق بالوقت + زر الإكسيل ---
        csv_data = f"تقرير غياب يوم: {today_iso}\\nالشركة المشغلة,الوردية,اسم الموظف,رقم الهوية,الوظيفة,القسم\\n"
        
        absent_html = f"<div class='grand-summary' style='border-color: var(--danger); text-align: center;'>"
        absent_html += f"<h2 style='color: var(--danger);'>🚨 سجل الغياب - {active_shift_title}</h2>"
        absent_html += f"<button class='export-btn' onclick='downloadExcel()'>📥 تحميل كشف الغياب (Excel)</button><div style='text-align: right;'>"
        
        has_any_shift_in_companies = False

        for company in df['operatorCompanyName'].unique():
            if pd.isna(company) or company == 'غير محدد': continue
            c_df = df[df['operatorCompanyName'] == company]
            
            comp_absent_html = ""
            company_has_active_shift = False
            
            for shift in c_df['workShiftName'].unique():
                shift_clean = str(shift).replace('أ','ا').replace('إ','ا').replace('آ','ا').replace('ة','ه').strip().lower()
                is_active = any(kw in shift_clean for kw in shift_keywords)

                if not is_active:
                    continue

                company_has_active_shift = True
                has_any_shift_in_companies = True

                shift_df = c_df[c_df['workShiftName'] == shift]
                absent_rows = ""
                absent_count = 0
                
                for _, row in shift_df.iterrows():
                    eid = str(row.get('employeeCode', '')).strip().lower()
                    nid = str(row.get('nationalId', '')).replace('.0', '').strip().lower()
                    
                    if eid not in present_ids and nid not in present_ids:
                        absent_count += 1
                        c_name = row.get('clean_name', '')
                        c_job = row.get('occupationName', '')
                        c_dept = row.get('clean_dept', '')
                        
                        absent_rows += f"<tr><td>{c_name}</td><td>{nid}</td><td>{c_job}</td><td>{c_dept}</td></tr>"
                        # إضافة الموظف لملف الإكسيل
                        csv_data += f"{company},{shift},{c_name},{nid},{c_job},{c_dept}\\n"
                
                if absent_count > 0:
                    comp_absent_html += f"""
                    <h3 style='color: var(--secondary); margin-top: 25px; border-right: 4px solid var(--danger); padding-right: 10px;'>📍 {shift} <span style='color: white; background: var(--danger); padding: 3px 10px; border-radius: 10px; font-size: 0.8em; margin-right: 10px;'>إجمالي الغياب: {absent_count}</span></h3>
                    <table class='absent-table'>
                        <tr><th>الاسم</th><th>الهوية</th><th>الوظيفة</th><th>القسم</th></tr>
                        {absent_rows}
                    </table>
                    """
                else:
                    comp_absent_html += f"""
                    <h3 style='color: var(--secondary); margin-top: 25px; border-right: 4px solid var(--success); padding-right: 10px;'>📍 {shift} <span style='color: white; background: var(--success); padding: 3px 10px; border-radius: 10px; font-size: 0.8em; margin-right: 10px;'>الغياب: صفر (0) ✅</span></h3>
                    <div class='zero-absent-msg'>اكتمل حضور جميع الموظفين المجدولين في هذه الوردية.</div>
                    """
            
            if company_has_active_shift:
                absent_html += f"<div class='company-card' style='box-shadow: 0 4px 15px rgba(0,0,0,0.05); padding: 25px;'><div class='company-title' style='color: var(--primary); font-size: 1.5em; padding-bottom: 10px;'>🏢 {company}</div>{comp_absent_html}</div>"

        if not has_any_shift_in_companies:
            absent_html += f"<h3 style='color:#7f8c8d; text-align:center; margin-top: 30px;'>لا توجد بيانات مسجلة لـ ({active_shift_title}) في جميع الشركات.</h3>"

        absent_html += "</div></div>"
        html_content += absent_html
        
        # --- سكربت التنزيل ---
        html_content += f"""
            <script>
            function downloadExcel() {{
                var csv = `{csv_data}`;
                var blob = new Blob(["\\uFEFF" + csv], {{ type: 'text/csv;charset=utf-8;' }});
                var url = URL.createObjectURL(blob);
                var link = document.createElement("a");
                link.href = url;
                link.download = "سجل_الغياب_{today_iso}.csv";
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }}
            </script>
            <div class="footer" dir="ltr">
                AUTOMATED SYSTEM PREPARED BY<br>
                <b style="color: var(--primary); font-size: 1.5em;">Eng. Abdulaziz Alshehri</b>
            </div>
        </body></html>
        """

        with open(REPORT_PATH, "w", encoding="utf-8") as f:
            f.write(html_content)

    except Exception as e:
        print(f"❌ حدث خطأ فني: {e}")

if __name__ == "__main__":
    generate_master_dashboard()
