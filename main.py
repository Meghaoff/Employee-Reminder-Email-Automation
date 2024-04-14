import smtplib
import openpyxl, os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
load_dotenv()

# Function to read employee data from Excel file
def read_employee_data(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    employee_data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        employee = {
            'name': row[0],
            'email': row[1],
            'reminder_date': row[2]
        }
        employee_data.append(employee)
    print(employee_data)
    return employee_data

# Function to send reminder email to employees
def send_reminder_email(employee):
    sender_email = os.getenv('SENDER_EMAIL') 
    print(sender_email)
    password = os.getenv('SENDER_PASSWORD')

    # Create message container
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = employee['email']
    msg['Subject'] = 'Reminder: Upcoming Event'

    # Email body
    body = f"Hi {employee['name']},\n\nThis is a reminder that the upcoming event is scheduled for {employee['reminder_date']}.\n\nBest regards,\nYour Company"

    # Attach body to message
    msg.attach(MIMEText(body, 'plain'))

    # Send the message via SMTP server
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
# Main function
def main():
    file_path = r'./Employee.xlsx'  # path of Excel file
    employee_data = read_employee_data(file_path)
    
    for employee in employee_data:
        send_reminder_email(employee)
        print(f"Reminder email sent to {employee['email']}")

if __name__ == "__main__":
    main()