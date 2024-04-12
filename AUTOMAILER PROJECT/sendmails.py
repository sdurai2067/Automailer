import tkinter as tk  # Importing the tkinter module for GUI
from tkinter import filedialog, messagebox  # Importing specific components from tkinter
from functools import partial  # Importing partial function from functools
import os  # Importing os module for operating system functionalities
import csv  # Importing csv module for CSV file handling
import smtplib  # Importing smtplib for sending emails
from email.mime.multipart import MIMEMultipart  # Importing MIMEMultipart for email composition
from email.mime.text import MIMEText  # Importing MIMEText for email text
from email.mime.application import MIMEApplication  # Importing MIMEApplication for attaching files

MY_ADDRESS = 'customemails2024@gmail.com'  # Email address
PASSWORD = 'wuue vpfg kbbd qqtx'  # Email password

# Function to send an email
def send_email(to_email, subject, body, attachments=[]):
    msg = MIMEMultipart()  # Create MIMEMultipart object
    msg['From'] = MY_ADDRESS  # Set sender email
    msg['To'] = to_email  # Set recipient email
    msg['Subject'] = subject  # Set email subject
    msg.attach(MIMEText(body, 'plain'))  # Attach email body as plain text
    for attachment in attachments:  # Iterate through attachments
        with open(attachment, 'rb') as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(attachment))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
            msg.attach(part)  # Attach each file
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(MY_ADDRESS, PASSWORD)
        server.send_message(msg)  # Send email

# Function to generate email template
def generate_template(template_file, **kwargs):
    with open(template_file, 'r') as template_file:
        template = template_file.read()  # Read template from file
    for key, value in kwargs.items():  # Replace placeholders with provided data
        template = template.replace('${' + key + '}', value)
    return template

# Main function to handle email sending
def main(template_file, csv_file, use_case):
    status_count = {'success': 0, 'failure': 0}  # Track success and failure counts
    with open(csv_file, 'r') as file:
        reader = csv.DictReader(file)  # Read CSV file
        for row in reader:
            name = row['NAME']  # Extract recipient name and email
            email = row['EMAILID']
            if use_case == 'hr_work':  # Handle hr_work use case
                position = row['POSITION']
                start_date = row['START_DATE']
                body = generate_template(template_file, NAME=name, POSITION=position, START_DATE=start_date)
                subject = f"JOB INTERVIEW-REGARDING"
                attachments = []  # Initialize attachments list
            elif use_case == 'marketing':  # Handle marketing use case
                product_name = row['PRODUCT_NAME']
                product_benefit = row['PRODUCT_BENEFIT']
                product_link = row['PRODUCT_LINK']
                body = generate_template(template_file, NAME=name, PRODUCT_NAME=product_name, PRODUCT_BENEFIT=product_benefit, PRODUCT_LINK=product_link)
                subject = f"NEW PRODUCT ANNOUNCEMENT REGARDING"
                attachments = []  # Initialize attachments list
                if 'PRODUCT_CATLOGUE' in row and row['PRODUCT_CATLOGUE']:  # Check for product catalogue attachment
                    attachments.append(row['PRODUCT_CATLOGUE'])  # Add product catalogue to attachments
            elif use_case == 'students':  # Handle students use case
                dm_score = row['DM']
                ns_score = row['NS']
                cc_score = row['CC']
                esia_score = row['ESIA']
                eai_score = row['EAI']
                total_score = row['TOTAL']
                body = generate_template(template_file, NAME=name, DM=dm_score, NS=ns_score, CC=cc_score, ESIA=esia_score, EAI=eai_score, TOTAL=total_score)
                subject = f"EXAMINATION REPORT-REGARDING"
                attachments = []  # No attachments for students
            try:
                send_email(email, subject, body, attachments)  # Send email
                status_count['success'] += 1  # Increment success count
                print(f"Email sent to {name} at {email}")  # Print success message
            except Exception as e:
                status_count['failure'] += 1  # Increment failure count
                print(f"Failed to send email to {name} at {email}: {e}")  # Print failure message
    print(f"\nEmail sending summary:")  # Print summary
    print(f"Successful: {status_count['success']}")
    print(f"Failed: {status_count['failure']}")

# Function to browse for file
def browse_file(entry):
    filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select File")
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# Function to send emails
def send_emails(template_file, csv_file, use_case):
    try:
        main(template_file, csv_file, use_case)  # Call main function
        messagebox.showinfo("Success", "Emails sent successfully!")  # Show success message
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send emails: {str(e)}")  # Show error message

# Function to create GUI
def create_gui():
    window = tk.Tk()  # Create tkinter window
    window.title("AUTOEMAILER")  # Set window title
    window.configure(bg="lightblue")  # Set window background color

    # Styling for the project name label
    project_name_font = ("cursive", 20, "bold")

    # Project name label
    project_name_label = tk.Label(window, text="AUTOEMAILER", font=project_name_font, bg="lightblue")
    project_name_label.grid(row=0, column=0, columnspan=3, pady=10)

    # Styling for other widgets
    widget_border_width = 3
    widget_border_color = "black"

    # Template file path label and entry
    tk.Label(window, text="Template File Path:", bg="lightblue").grid(row=1, column=0, padx=5, pady=5)
    template_entry = tk.Entry(window, width=50, bd=widget_border_width, relief="solid")
    template_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(window, text="Browse", command=partial(browse_file, template_entry), bd=widget_border_width, relief="solid").grid(row=1, column=2, padx=5, pady=5)

    # CSV file path label and entry
    tk.Label(window, text="CSV File Path:", bg="lightblue").grid(row=2, column=0, padx=5, pady=5)
    csv_entry = tk.Entry(window, width=50, bd=widget_border_width, relief="solid")
    csv_entry.grid(row=2, column=1, padx=5, pady=5)
    tk.Button(window, text="Browse", command=partial(browse_file, csv_entry), bd=widget_border_width, relief="solid").grid(row=2, column=2, padx=5, pady=5)

    # Use Case label and dropdown
    tk.Label(window, text="Use Case:", bg="lightblue").grid(row=3, column=0, padx=5, pady=5)
    use_case_var = tk.StringVar(window)
    use_case_var.set("hr_work")  # Default value
    use_case_menu = tk.OptionMenu(window, use_case_var, "hr_work", "marketing", "students")
    use_case_menu.config(bd=widget_border_width, relief="solid")
    use_case_menu.grid(row=3, column=1, padx=5, pady=5)

    # Send Emails button
    send_button = tk.Button(window, text="Send Emails", command=lambda: send_emails(template_entry.get(), csv_entry.get(), use_case_var.get()), bd=widget_border_width, relief="solid")
    send_button.grid(row=4, column=1, padx=5, pady=5)

    window.mainloop()

if __name__ == "__main__":
    create_gui()

