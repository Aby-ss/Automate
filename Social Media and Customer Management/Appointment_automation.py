import smtplib
import schedule
import time

def send_email(recipient_email, subject, message):
    # SMTP server configuration (update with your own values)
    smtp_server = 'smtp.example.com'
    smtp_port = 587
    smtp_username = 'your_username'
    smtp_password = 'your_password'

    # Create SMTP connection
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)

        # Construct email message
        email_message = f"Subject: {subject}\n\n{message}"

        # Send email
        server.sendmail(smtp_username, recipient_email, email_message)

def set_appointment_reminder(recipient_email, appointment_datetime):
    subject = 'Appointment Reminder'
    message = f"Hi there!\n\nThis is a reminder for your appointment on {appointment_datetime}.\n\nPlease be prepared for the call.\n\nBest regards,\nYour Company"

    # Schedule the email to be sent at the specified appointment time
    schedule.every().day.at(appointment_datetime).do(send_email, recipient_email, subject, message)

    print(f"Appointment reminder set for {appointment_datetime}")

# Example usage
recipient_email = 'client@example.com'
appointment_datetime = '2023-07-16 09:00'  # Set the appointment date and time in YYYY-MM-DD HH:MM format

set_appointment_reminder(recipient_email, appointment_datetime)

# Keep the program running to execute scheduled tasks
while True:
    schedule.run_pending()
    time.sleep(1)
  
