import pandas as pd
import smtplib
import os
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime


# ---------------------------------------------------------------
# Email settings
# Put your email credentials in environment variables, never
# hardcode them directly in the script.
#
# How to set them (run once in your terminal):
#   Windows:
#       set EMAIL_ADDRESS=your_email@gmail.com
#       set EMAIL_PASSWORD=your_app_password
#   Mac/Linux:
#       export EMAIL_ADDRESS=your_email@gmail.com
#       export EMAIL_PASSWORD=your_app_password
#
# Gmail users: create an App Password from your Google account
# security settings and use it instead of your real password.
# ---------------------------------------------------------------

SENDER_EMAIL   = os.environ.get("EMAIL_ADDRESS")
SENDER_PASSWORD = os.environ.get("EMAIL_PASSWORD")
SMTP_SERVER    = "smtp.gmail.com"
#SMTP_PORT      = 465
SMTP_PORT = 587

# Path to the Excel file that the sales manager updates
EXCEL_FILE = r".\automation\orders_sample.xlsx"

# How often the script checks for new orders (in seconds)
CHECK_INTERVAL = 30


def load_orders(file_path):
    # Read the Excel file and return a DataFrame
    # Expected columns: order_id, customer_name, customer_email, product, quantity, date
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return pd.DataFrame()

    df = pd.read_excel(file_path, dtype={"order_id": str})
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df


def send_confirmation_email(order):
    # Build and send a confirmation email to the customer
    receiver_email = order["customer_email"]
    customer_name  = order["customer_name"]
    order_id       = order["order_id"]
    product        = order["product"]
    quantity       = order["quantity"]
    order_date     = order.get("date", datetime.today().strftime("%Y-%m-%d"))

    subject = f"Order Confirmation - Order #{order_id}"

    # Plain text email body
    body = f"""Dear {customer_name},

Thank you for your order. We have successfully received it and it is now being processed.

Order Details:
--------------
Order ID  : {order_id}
Product   : {product}
Quantity  : {quantity}
Date      : {order_date}

We will notify you once your order has been shipped.

Best regards,
Sales Team
"""

    # Compose the email message
    message = MIMEMultipart()
    message["From"]    = SENDER_EMAIL
    message["To"]      = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    # Connect to the SMTP server and send the email
    
    
    
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:    
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, receiver_email, message.as_string())
        print(f"Confirmation sent to {receiver_email} for order #{order_id}")
        return True
    except Exception as e:
        print(f"Failed to send email for order #{order_id}: {e}")
        return False


def load_sent_orders(log_file):
    # Load the list of order IDs that already received a confirmation email
    if not os.path.exists(log_file):
        return set()

    with open(log_file, "r") as f:
        return set(line.strip() for line in f if line.strip())


def save_sent_order(log_file, order_id):
    # Append a new order ID to the sent log so it is not emailed again
    with open(log_file, "a") as f:
        f.write(str(order_id) + "\n")


def watch_excel_and_notify(excel_file, log_file="sent_orders.txt"):
    # Main loop: check the Excel file every CHECK_INTERVAL seconds
    # and send an email for any new order that has not been notified yet

    print("Watching for new orders...")
    print(f"File : {excel_file}")
    print(f"Check interval: {CHECK_INTERVAL} seconds")
    print("-" * 40)

    # Validate email credentials before starting
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("Error: EMAIL_ADDRESS or EMAIL_PASSWORD environment variable is not set.")
        return

    while True:
        orders = load_orders(excel_file)

        if orders.empty:
            print("No orders found or file could not be read.")
        else:
            sent_orders = load_sent_orders(log_file)

            # Process only orders that have not been notified yet
            new_orders = orders[~orders["order_id"].isin(sent_orders)]

            if new_orders.empty:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] No new orders.")
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Found {len(new_orders)} new order(s).")

                for _, order in new_orders.iterrows():
                    success = send_confirmation_email(order)
                    if success:
                        save_sent_order(log_file, order["order_id"])

        # Wait before checking again
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    watch_excel_and_notify(EXCEL_FILE)
