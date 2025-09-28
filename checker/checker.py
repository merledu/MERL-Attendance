#!/usr/bin/env python3
"""
Read today's attendance from Google Sheets.

- Finds today's IN/OUT columns.
- Returns dictionary with present students + IN times, and absent students.
"""

import datetime as dt
from google.oauth2.service_account import Credentials
import gspread
from icecream import ic

# ---------------------- CONFIG ----------------------
SERVICE_ACCOUNT_FILE = "attendance.json"
SPREADSHEET_ID = "1NnAM6FtuVrbf3h7dbIwCRHUj4UJ2wDO30_jmq1leGkM"
MASTER_STUDENT_SHEET_TITLE = "Students"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ----------------------------------------------------

def get_today_attendance():
    today_str = str(dt.date.today())
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    ss = client.open_by_key(SPREADSHEET_ID)

    # open current month worksheet
    month_title = dt.date.today().strftime("%B %Y")
    try:
        ws = ss.worksheet(month_title)
    except Exception as e:
        print(f"No worksheet found for {month_title}: {e}")
        return {"present": {}, "absent": []}

    # find today's columns
    row1 = ws.row_values(1)
    date_cols = {}
    col = 2
    while col <= len(row1):
        if row1[col - 1]:
            date_cols[row1[col - 1].strip()] = (col, col + 1)
            col += 2
        else:
            break

    if today_str not in date_cols:
        print(f"No attendance columns yet for {today_str}")
        return {"present": {}, "absent": []}

    in_col, out_col = date_cols[today_str]

    # read all student names and in-times
    names = ws.col_values(1)[2:]  # students start from row 3
    in_times = ws.col_values(in_col)[2:]

    ic(names)
    ic(in_times)

    present = {}
    absent = []

    for i in range(len(names)):
        name = names[i].strip()
        if not name:
            continue
        intime = in_times[i].strip() if i < len(in_times) else ""
        if intime:
            present[name] = intime
        else:
            absent.append(name)


    return {"present": present, "absent": absent}

import requests
import json
from datetime import datetime

def send_attendance_report(api_base_url, recipients, present_employees_dict, absent_employees_list):
    """
    Send creative attendance report email via your API route
    
    Parameters:
    - api_base_url: Base URL of your API (e.g., "http://localhost:5000")
    - recipients: List of recipient email addresses
    - present_employees_dict: Dictionary with employee names as keys and in-time as values
    - absent_employees_list: List of absent employee names
    """
    
    # Calculate summary statistics
    present_count = len(present_employees_dict)
    absent_count = len(absent_employees_list)
    
    # Count employees who arrived on time (before 10:00 AM)
    on_time_count = 0
    
    for name, time_str in present_employees_dict.items():
        # Parse time and check if before 9:15
        try:
            time_obj = datetime.strptime(time_str, "%H:%M").time()
            if time_obj <= datetime.strptime("10:00", "%H:%M").time():
                on_time_count += 1
        except:
            pass
    
    # Get current date
    current_date = datetime.now().strftime("%A, %B %d, %Y")
    
    # Generate the HTML email content
    html_content = generate_email_html(
        current_date=current_date,
        present_count=present_count,
        absent_count=absent_count,
        on_time_count=on_time_count,
        present_employees_dict=present_employees_dict,
        absent_employees_list=absent_employees_list
    )
    
    # Prepare the email data
    email_data = {
        "subject": f"MERL Attendance Report - {datetime.now().strftime('%Y-%m-%d')}",
        "recipients": recipients,
        "body": html_content
    }
    
    # Send request to your API route
    try:
        response = requests.post(
            f"{api_base_url}/api/send-email",
            headers={"Content-Type": "application/json"},
            data=json.dumps(email_data)
        )
        
        if response.status_code == 202:
            result = response.json()
            print(f"✅ Email scheduled successfully! Task ID: {result.get('task_id')}")
            return True
        else:
            print(f"❌ Failed to schedule email. Status: {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"❌ Error connecting to API: {e}")
        return False
def generate_email_html(current_date, present_count, absent_count, on_time_count, 
                       present_employees_dict, absent_employees_list):
    """Generate the HTML email content with the new template"""
    
    # Generate present employees table rows
    present_rows = ""
    for name, time in present_employees_dict.items():
        # Determine time class based on arrival time
        time_class = "time-early"
        try:
            time_obj = datetime.strptime(time, "%H:%M").time()
            if time_obj > datetime.strptime("09:30", "%H:%M").time():
                time_class = "time-very-late"
            elif time_obj > datetime.strptime("09:15", "%H:%M").time():
                time_class = "time-late"
        except:
            pass
            
        present_rows += f"""
        <tr>
            <td><strong>{name.title()}</strong></td>
            <td class="{time_class}">{time}</td>
            <td><span class="status-badge status-present">Present</span></td>
        </tr>
        """
    
    # Generate absent employees table rows
    absent_rows = ""
    for name in absent_employees_list:
        absent_rows += f"""
        <tr>
            <td><strong>{name.title()}</strong></td>
            <td><span class="status-badge status-absent">Absent</span></td>
        </tr>
        """
    
    # Read the HTML template file
    with open("attendance.html", "r") as file:
        html_template = file.read()
    
    # Replace placeholders
    html_content = html_template.replace("[DATE_PLACEHOLDER]", current_date)
    html_content = html_content.replace("[PRESENT_COUNT]", str(present_count))
    html_content = html_content.replace("[ABSENT_COUNT]", str(absent_count))
    html_content = html_content.replace("[ON_TIME_COUNT]", str(on_time_count))
    html_content = html_content.replace("[PRESENT_DATA_PLACEHOLDER]", present_rows)
    html_content = html_content.replace("[ABSENT_DATA_PLACEHOLDER]", absent_rows)
    
    return html_content

if __name__ == "__main__":
    data = get_today_attendance()
    print("Today's Attendance:")
    print(data)

    # from checker.send_mail import send_attendance_report
    send_attendance_report(
        "http://116.213.35.22:8016",
        ["shahzaibceo@gmail.com"],
        data["present"], data["absent"])
