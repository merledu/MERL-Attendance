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
    
    # Count employees who arrived on time (before 9:15 AM)
    on_time_count = 0
    
    for name, time_str in present_employees_dict.items():
        # Parse time and check if before 9:15
        try:
            time_obj = datetime.strptime(time_str, "%H:%M").time()
            if time_obj <= datetime.strptime("09:15", "%H:%M").time():
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
    """Generate the HTML email content"""
    
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
            <td>{name}</td>
            <td class="{time_class}">{time}</td>
            <td class="status-present">Present</td>
        </tr>
        """
    
    # Generate absent employees table rows
    absent_rows = ""
    for name in absent_employees_list:
        absent_rows += f"""
        <tr>
            <td>{name}</td>
            <td class="status-absent">Absent</td>
        </tr>
        """
    
    # The complete HTML template
    html_template = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>MERL Daily Attendance Report</title>
        <style>
            /* Reset margins and padding */
            body, html {{
                margin: 0 !important;
                padding: 0 !important;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f5f7fa;
                color: #333;
                line-height: 1.4;
            }}
            
            .email-container {{
                width: 100%;
                max-width: 600px;
                margin: 0 auto;
                background-color: #ffffff;
                border-radius: 12px;
                overflow: hidden;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            }}
            
            /* Header with no top spacing */
            .header {{
                background: linear-gradient(135deg, #4b6cb7 0%, #182848 100%);
                color: white;
                padding: 25px 30px;
                text-align: center;
                margin: 0;
            }}
            
            .header h1 {{
                margin: 0;
                font-size: 24px;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
            }}
            
            .header-icon {{
                font-size: 28px;
            }}
            
            .date-display {{
                margin-top: 10px;
                font-size: 16px;
                opacity: 0.9;
            }}
            
            .content {{
                padding: 25px 30px;
                margin: 0;
            }}
            
            .section-title {{
                font-size: 18px;
                color: #4b6cb7;
                margin-bottom: 15px;
                padding-bottom: 8px;
                border-bottom: 2px solid #eaeaea;
                display: flex;
                align-items: center;
                gap: 8px;
            }}
            
            .section-icon {{
                font-size: 20px;
            }}
            
            /* Summary Cards - fixed width issues */
            .summary-cards {{
                display: flex;
                justify-content: space-between;
                margin-bottom: 25px;
                gap: 10px;
                width: 100%;
            }}
            
            .summary-card {{
                flex: 1;
                min-width: 0;
                background: linear-gradient(135deg, #f0f4ff 0%, #e0e7ff 100%);
                border-radius: 8px;
                padding: 15px 10px;
                text-align: center;
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);
                box-sizing: border-box;
            }}
            
            .summary-number {{
                font-size: 24px;
                font-weight: 700;
                color: #4b6cb7;
                margin: 5px 0;
                line-height: 1.2;
            }}
            
            .summary-label {{
                font-size: 12px;
                color: #666;
                line-height: 1.2;
            }}
            
            /* Table Styles */
            .attendance-table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 25px;
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
            }}
            
            .attendance-table th {{
                background-color: #f0f4ff;
                padding: 12px 15px;
                text-align: left;
                font-weight: 600;
                color: #4b6cb7;
                border-bottom: 2px solid #e0e7ff;
            }}
            
            .attendance-table td {{
                padding: 12px 15px;
                border-bottom: 1px solid #f0f0f0;
            }}
            
            .attendance-table tr:hover {{
                background-color: #f9faff;
            }}
            
            .status-present {{
                color: #2ecc71;
                font-weight: 600;
            }}
            
            .status-absent {{
                color: #e74c3c;
                font-weight: 600;
            }}
            
            .time-early {{
                color: #27ae60;
            }}
            
            .time-late {{
                color: #e67e22;
            }}
            
            .time-very-late {{
                color: #c0392b;
            }}
            
            /* Footer */
            .footer {{
                background-color: #f9f9f9;
                padding: 20px 30px;
                text-align: center;
                border-top: 1px solid #eaeaea;
                font-size: 14px;
                color: #777;
            }}
            
            .bot-signature {{
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                margin-bottom: 10px;
                width: 100%;
            }}
            
            .bot-icon {{
                font-size: 20px;
                color: #4b6cb7;
            }}
            
            /* Responsive */
            @media (max-width: 600px) {{
                .summary-cards {{
                    flex-direction: column;
                }}
                
                .summary-card {{
                    margin-bottom: 10px;
                }}
                
                .attendance-table {{
                    font-size: 14px;
                }}
                
                .attendance-table th, 
                .attendance-table td {{
                    padding: 8px 10px;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="email-container">
            <div class="header">
                <h1><span class="header-icon">📊</span> MERL Attendance Report</h1>
                <div class="date-display">Today, {current_date}</div>
            </div>
            
            <div class="content">
                <div class="summary-cards">
                    <div class="summary-card">
                        <div class="summary-number">{present_count}</div>
                        <div class="summary-label">Present Today</div>
                    </div>
                    <div class="summary-card">
                        <div class="summary-number">{absent_count}</div>
                        <div class="summary-label">Absent Today</div>
                    </div>
                    <div class="summary-card">
                        <div class="summary-number">{on_time_count}</div>
                        <div class="summary-label">On Time</div>
                    </div>
                </div>
                
                <div class="section-title">
                    <span class="section-icon">✅</span> Present Employees
                </div>
                
                <table class="attendance-table">
                    <thead>
                        <tr>
                            <th>Employee Name</th>
                            <th>Arrival Time</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {present_rows}
                    </tbody>
                </table>
                
                <div class="section-title">
                    <span class="section-icon">❌</span> Absent Employees
                </div>
                
                <table class="attendance-table">
                    <thead>
                        <tr>
                            <th>Employee Name</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {absent_rows}
                    </tbody>
                </table>
                
                <div style="margin-top: 25px; padding: 15px; background-color: #f9f9f9; border-radius: 8px; font-size: 14px;">
                    <strong>Note:</strong> This report was automatically generated by the MERL Attendance Bot. 
                    Please contact HR if you notice any discrepancies.
                </div>
            </div>
            
            <div class="footer">
                <div>🤖 Sent by MERL Attendance Bot</div>
                <div>This is an automated message. Please do not reply to this email.</div>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html_template

# Example usage with dictionary data structure
if __name__ == "__main__":
    # Sample data with dictionary structure
    present_employees_dict = {
        "John Doe": "08:45",
        "Jane Smith": "09:05", 
        "Bob Johnson": "09:25",
        "Alice Williams": "08:30",
        "Michael Brown": "09:40"
    }
    
    absent_employees_list = [
        "Charlie Brown",
        "Diana Prince"
    ]
    
    # Send the report
    send_attendance_report(
        api_base_url="http://localhost:5000",
        recipients=["manager@merl.com", "hr@merl.com"],
        present_employees_dict=present_employees_dict,
        absent_employees_list=absent_employees_list
    )