# Import necessary libraries
import tkinter as tk
from tkinter import filedialog, messagebox
from functools import partial
import os
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Define email credentials
MY_ADDRESS = 'customemails2024@gmail.com'
PASSWORD = 'wuue vpfg kbbd qqtx'

# Function to send email
def send_email(to_email, subject, body, attachments=[]):
    msg = MIMEMultipart()
    msg['From'] = MY_ADDRESS
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    for attachment in attachments:
        with open(attachment, 'rb') as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(attachment))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
            msg.attach(part)
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(MY_ADDRESS, PASSWORD)
        server.send_message(msg)

# Function to generate email template
def generate_template(template_file, download_link=None, **kwargs):
    with open(template_file, 'r') as template_file:
        template = template_file.read()
    for key, value in kwargs.items():
        template = template.replace('${' + key + '}', value)
    if download_link:
        template += f"\n\nHere is your email attachment: <a href='{download_link}'>Download</a>"
    return template

# Main function to send emails based on CSV data and use case
def main(template_file, csv_file, use_case):
    status_count = {'success': 0, 'failure': 0}
    with open(csv_file, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            name = row['NAME']
            email = row['EMAILID']
            if use_case == 'hr_work':
                position = row['POSITION']
                start_date = row['START_DATE']
                body = generate_template(template_file, NAME=name, POSITION=position, START_DATE=start_date)
                subject = f"JOB INTERVIEW-REGARDING"
                attachments = [row['CATLOUGE']] if 'CATLOUGE' in row else []
            elif use_case == 'marketing':
                product_name = row['PRODUCT_NAME']
                product_benefit = row['PRODUCT_BENEFIT']
                product_link = row['PRODUCT_LINK']
                body = generate_template(template_file, NAME=name, PRODUCT_NAME=product_name, PRODUCT_BENEFIT=product_benefit, PRODUCT_LINK=product_link)
                subject = f"NEW PRODUCT ANNOUNCEMENT REGARDING"
                attachments = [row['PRODUCT_CATLOGUE']] if 'PRODUCT_CATLOGUE' in row else []
            elif use_case == 'students':
                dm_score = row['DM']
                ns_score = row['NS']
                cc_score = row['CC']
                esia_score = row['ESIA']
                eai_score = row['EAI']
                total_score = row['TOTAL']
                body = generate_template(template_file,NAME=name,DM=dm_score,NS=ns_score,CC=cc_score,ESIA=esia_score,EAI=eai_score,TOTAL=total_score)
                subject = f"EXAMINATION REPORT-REGARDING"
                attachments = []
            try:
                send_email(email, subject, body, attachments)
                status_count['success'] += 1
                print(f"Email sent to {name} at {email}")
            except Exception as e:
                status_count['failure'] += 1
                print(f"Failed to send email to {name} at {email}: {e}")
    print(f"\nEmail sending summary:")
    print(f"Successful: {status_count['success']}")
    print(f"Failed: {status_count['failure']}")

# Function to browse file and insert path into entry field
def browse_file(entry):
    filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select File")
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# Function to initiate email sending process
def send_emails(template_file, csv_file, use_case):
    try:
        main(template_file, csv_file, use_case)
        messagebox.showinfo("Success", "Emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send emails: {str(e)}")

# Function to create GUI
def create_gui():
    window = tk.Tk()
    window.title("AUTOEMAILER")
    window.configure(bg="lightblue")

    project_name_font = ("cursive", 20, "bold")
# Set GUI background color to light blue
    project_name_label = tk.Label(window, text="AUTOEMAILER", font=project_name_font,bg="lightblue")
    project_name_label.grid(row=0, column=0, columnspan=3, pady=10)

    widget_border_width = 3

    # Label and entry for Template File Path
    tk.Label(window, text="Template File Path:", bg="lightblue").grid(row=1, column=0, padx=5, pady=5)
    template_entry = tk.Entry(window, width=50, bd=widget_border_width, relief="solid")
    template_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(window, text="Browse", command=partial(browse_file, template_entry), bd=widget_border_width, relief="solid").grid(row=1, column=2, padx=5, pady=5)

    # Label and entry for CSV File Path
    tk.Label(window, text="CSV File Path:", bg="lightblue").grid(row=2, column=0, padx=5, pady=5)
    csv_entry = tk.Entry(window, width=50, bd=widget_border_width, relief="solid")
    csv_entry.grid(row=2, column=1, padx=5, pady=5)
    tk.Button(window, text="Browse", command=partial(browse_file, csv_entry), bd=widget_border_width, relief="solid").grid(row=2, column=2, padx=5, pady=5)

    # Label and dropdown for selecting Use Case
    tk.Label(window, text="Use Case:", bg="lightblue").grid(row=3, column=0, padx=5, pady=5)
    use_case_var = tk.StringVar(window)
    use_case_var.set("hr_work")
    use_case_menu = tk.OptionMenu(window, use_case_var, "hr_work", "marketing", "students")
    use_case_menu.config(bd=widget_border_width, relief="solid")
    use_case_menu.grid(row=3, column=1, padx=5, pady=5)

    # Button to initiate sending emails
    send_button = tk.Button(window, text="Send Emails", command=lambda: send_emails(template_entry.get(), csv_entry.get(), use_case_var.get()), bd=widget_border_width, relief="solid")
    send_button.grid(row=4, column=1, padx=5, pady=5)

    window.mainloop()

if __name__ == "__main__":
    create_gui()
