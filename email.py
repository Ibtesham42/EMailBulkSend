import streamlit as st
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import io
import plotly.express as px

# App title
st.title("Bulk Email Sender with Tracking and Statistics")

# Initialize session state for email logs
if "email_log" not in st.session_state:
    st.session_state.email_log = []

# Section: Email Configuration
st.header("Email Configuration")
sender_email = st.text_input("Sender Email", placeholder="Enter your email address")
password = st.text_input("Password", type="password", placeholder="Enter your email password")
subject = st.text_input("Email Subject", placeholder="Enter the subject of your email")

# Section: Upload Excel File
st.header("Upload Recipient Details")
uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "csv"], help="Upload a file with columns like 'Email', 'First name', 'Domain'.")

# Section: Write Email Body
st.header("Email Content")
body_template = st.text_area(
    "Write Email Body",
    value="""Hello {fname},

Thank you for applying for the internship position in {domain}. We look forward to discussing this opportunity further.

Warm regards,
HR Team""",
    help="Use placeholders like {fname} and {domain} to personalize the email.",
)

# Section: Attachment (Optional)
st.header("Attachment (Optional)")
attachment_file = st.file_uploader("Upload an attachment", type=["pdf", "txt", "docx", "xlsx"], help="Optional: Attach a file to the email.")

# Send Emails Button
if st.button("Send Emails"):
    if not sender_email or not password or not subject or not uploaded_file:
        st.error("Please fill in all required fields!")
    else:
        try:
            # Read uploaded file
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            # Check for required columns
            required_columns = ['Email', 'First name', 'Domain']
            if not all(col in df.columns for col in required_columns):
                st.error(f"Uploaded file must contain the following columns: {', '.join(required_columns)}")
            else:
                # Process each recipient
                successes, failures = 0, 0
                for _, row in df.iterrows():
                    receiver_email = row['Email']
                    fname = row['First name']
                    domain = row['Domain']
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Customize the email body
                    body = body_template.format(fname=fname, domain=domain)

                    # Create the email
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = receiver_email
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'plain'))

                    # Add attachment (if uploaded)
                    if attachment_file:
                        try:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment_file.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename={attachment_file.name}')
                            msg.attach(part)
                        except Exception as e:
                            st.warning(f"Error attaching file for {receiver_email}: {e}")

                    # Send the email
                    try:
                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.starttls()
                        server.login(sender_email, password)
                        server.sendmail(sender_email, receiver_email, msg.as_string())
                        st.session_state.email_log.append({"Email": receiver_email, "Name": fname, "Domain": domain, "Status": "Success", "Timestamp": timestamp})
                        successes += 1
                    except Exception as e:
                        st.session_state.email_log.append({"Email": receiver_email, "Name": fname, "Domain": domain, "Status": f"Failed ({e})", "Timestamp": timestamp})
                        failures += 1
                    finally:
                        server.quit()

                # Summary
                st.success(f"Emails sent successfully to {successes} recipients!")
                if failures:
                    st.warning(f"Failed to send emails to {failures} recipients.")

        except Exception as e:
            st.error(f"An error occurred: {e}")

# Section: Download Log
st.header("Download Email Log")
if st.session_state.email_log:
    log_df = pd.DataFrame(st.session_state.email_log)
    st.write(log_df)
    csv = log_df.to_csv(index=False)
    st.download_button(label="Download Log as CSV", data=csv, file_name="email_log.csv", mime="text/csv")

    # Statistics
    st.header("Statistics")
    status_counts = log_df['Status'].value_counts()
    fig = px.pie(values=status_counts, names=status_counts.index, title="Email Sending Status")
    st.plotly_chart(fig)
else:
    st.info("No email logs available yet.")
