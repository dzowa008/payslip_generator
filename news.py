import pandas as pd
import smtplib
import os
from fpdf import FPDF
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()
employee_data = pd.read_excel('Book3.xlsx', engine='openpyxl')
employee_data.columns = employee_data.columns.str.strip()

required_columns = ['Employee ID', 'Name', 'Email', 'Basic Salary', 'Allowance', 'Deductions']
for col in required_columns:
    if col not in employee_data.columns:
        employee_data[col] = 0  

# Convert salary columns to numeric values
employee_data['Basic Salary'] = pd.to_numeric(employee_data['Basic Salary'], errors='coerce').fillna(0)
employee_data['Allowance'] = pd.to_numeric(employee_data['Allowance'], errors='coerce').fillna(0)
employee_data['Deductions'] = pd.to_numeric(employee_data['Deductions'], errors='coerce').fillna(0)

class PayslipPDF(FPDF):
    def header(self):
        """Header Section - Adds Company Title"""
        self.set_font("Arial", "B", 18)
        self.set_text_color(0, 51, 102)
        self.cell(200, 10, "Boss_smarta.zw", ln=True, align='C')
        self.set_font("Arial", "I", 12)
        self.cell(200, 10, "Payslip for the Month", ln=True, align='C')
        self.ln(5)
        self.set_text_color(0, 0, 255) 
        self.cell(200, 2, "=" * 100, ln=True)
        self.ln(5)

    def footer(self):
        """Footer Section - Adds Page Number"""
        self.set_y(-15)
        self.set_font("Arial", "I", 10)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Page {self.page_no()}", align='C')


SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS") 
APP_PASSWORD =  os.getenv("EMAIL_PASSWORD")

def send_email(to_email, subject, body, attachment_path):
    """Function to send email with an attachment."""
    msg = EmailMessage()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    
    with open(attachment_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(attachment_path))

    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, APP_PASSWORD)
            server.send_message(msg)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {e}")


for index, row in employee_data.iterrows():
    pdf = PayslipPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(40, 10, "Employee Details:", ln=True)
    pdf.set_text_color(0, 0, 0)  
    pdf.set_font("Arial", "", 12)
    pdf.cell(200, 8, f"Employee ID   : {row['Employee ID']}", ln=True)
    pdf.cell(200, 8, f"Name          : {row['Name']}", ln=True)
    pdf.cell(200, 8, "-" * 200, ln=True)
    pdf.set_text_color(0, 0, 0)  
    pdf.set_font("Arial", "B", 12)
    pdf.cell(40, 10, "Salary Details:", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.cell(200, 8, f"Basic Salary  : $ {row['Basic Salary']:.2f}", ln=True)
    pdf.cell(200, 8, f"Allowance     : $ {row['Allowance']:.2f}", ln=True)
    pdf.cell(200, 8, f"Deductions    : $ {row['Deductions']:.2f}", ln=True)
    pdf.set_text_color(0, 0, 255) 

    pdf.cell(200, 2, "=" * 80, ln=True)
    pdf.set_text_color(0, 0,0)  
    net_salary = row['Basic Salary'] + row['Allowance'] - row['Deductions']
    pdf.set_font("Arial", "B", 12)
    pdf.cell(200, 10, f"Net Salary    : $ {net_salary:.2f}", ln=True, align='R')

    
    filename = f"{row['Name'].replace(' ', '_')}_Payslip.pdf"
    pdf.output(filename)

    
    if row.get("Email"): 
        subject = "Your Monthly Payslip"
        body = f"Dear {row['Name']},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nGWAVAVA ENTERPRISE"
        send_email(row["Email"], subject, body, filename)

print("All Payslips Successfully Created and Sent!")