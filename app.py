import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io

st.set_page_config(page_title="NSS Email Notification System")
# Streamlit app
st.title("Automated Email Notification System")

# User inputs for email credentials
outlook_user = st.text_input("Outlook Email Address", "")
outlook_password = st.text_input("Outlook Password", "", type="password")
# Define the email template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Venue Notification</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f8f9fa;
            color: #2c3e50;
            line-height: 1.6;
        }
        .email-container {
            max-width: 700px;
            margin: auto;
            padding: 20px;
            background: #fff;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        .header {
            background: #d9534f;
            padding: 10px;
            text-align: center;
            color: white;
            border-radius: 10px 10px 0 0;
        }
        .content {
            padding: 20px;
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 12px;
            color: gray;
        }
    </style>
</head>
<body>
    <div class="header">
    <img src="https://i.imgur.com/jkpf1qu.png" alt="NSS Logo" width="200" 
         style="display: block; margin: 0 auto; padding-bottom: 5px;">
    <h1 style="margin-top: 0;">Venue Notification</h1>
</div>

        <div class="content">
            <p>This is to inform you that the venue for the <b>[event_name]</b> has been updated to <b>[venue]</b>.</p>
            <p>Kindly ensure your presence at the updated venue by <b>[Time]</b>,<b>[Date]</b> sharp.</p>
            <p>We appreciate your cooperation and look forward to your enthusiastic participation.</p>
            <p>Best regards,<br>
            <b>[event_incharge_name]</b><br>
            Event In-Charge<br>
            NSS Unit</p>
        </div>
        <div class="footer">
            <p>For more details</p> 
            <p>Visit our Instagram: <a href="https://www.instagram.com/klef_nss_official/">@klef_nss_official</a></p>
            <p>Join our Telegram: <a href="https://t.me/+k_Bt9R_WDxVjNGJl">@KLEF_NSS_Y23 BATCH</a></p>
            <br>
        </div>
    </div>
</body>
</html>
"""
with st.sidebar:
    st.header("Instructions")
    st.write("""
    - Your file **must contain** the following columns:
      - `event_name`: Name of the event  
      - `venue`: Location of the event  
      - `Time`: Event start time  
      - `event_incharge_name`: Name of the event in-charge  
      - `student_id_number`: Student’s university ID  
      - `Date` : Date of the event
    - **Do not include extra columns or empty rows.**
    - The system will generate emails using `student_id_number@kluniversity.in`.
    """)
    # Create a sample dataframe
    sample_data = pd.DataFrame({
        "event_name": ["Tech Talk", "AI Workshop"],
        "venue": ["Indoor Stadium", "Lab 205"],
        "Time": ["1:30 PM", "2:00 PM"],
        "Date": ["20/02/2025","21/03/2025"],
        "event_incharge_name": ["Aravind", "Rahul"],
        "student_id_number": ["2200080234", "2300080456"]
    })

    # Convert to CSV
    csv_buffer = io.StringIO()
    sample_data.to_csv(csv_buffer, index=False)
    csv_data = csv_buffer.getvalue()

    # Streamlit UI
    st.markdown("### 📥 Download Sample CSV")
    st.write("Ensure your file follows this format:")
    st.dataframe(sample_data)

    # Add download button
    st.download_button(
        label="📂 Download Sample CSV",
        data=csv_data,
        file_name="sample_event_data.csv",
        mime="text/csv"
    )


# Upload CSV or Excel
uploaded_file = st.file_uploader(
    "Upload CSV or Excel file (with event_name, venue, Time, Date, event_incharge_name, student_id_number)",
    type=["csv", "xlsx", "xls"]
)

# File uploader for PDF or DOCX
attachment_file = st.file_uploader(
    "Upload a PDF, DOCX, JPEG, PNG, or other file for permissions as an attachment (optional)", 
    type=None,  # Allow all file types
    accept_multiple_files=True
)

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            data = pd.read_csv(uploaded_file)
        else:
            data = pd.read_excel(uploaded_file, engine="openpyxl")  # Ensure 'openpyxl' is installed for xlsx

        # Strip spaces from column names
        data.columns = data.columns.str.strip()

        # Drop rows where essential fields are NaN
        required_columns = ["event_name", "venue", "Time", "Date","event_incharge_name", "student_id_number"]
        data = data.dropna(subset=required_columns)

        data['student_id_number'] = data['student_id_number'].astype(str).str.strip()
        data['student_id_number'] = data['student_id_number'].apply(lambda x: x.split(".")[0] if "." in x else x)
        st.write("📂 Uploaded Data Preview:", data.head())

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
    emails = []

    for _, row in data.iterrows():
        email_body = html_template.replace("[event_name]", str(row.get("event_name"))) \
            .replace("[venue]", str(row.get("venue"))) \
            .replace("[Time]", str(row.get("Time"))) \
            .replace("[Date]", str(row.get("Date"))) \
            .replace("[event_incharge_name]", str(row.get("event_incharge_name")))
        student_email = f"{row.get('student_id_number')}@kluniversity.in"

        emails.append({
            "To": student_email,
            "Subject": f"Venue Update for {row.get('event_name', 'Unknown Event')} on {row.get('Date','unknown Date')}",
            "Body": email_body
        })

    emails_df = pd.DataFrame(emails)
    st.write(emails_df)

    # Button to send emails
    if st.button("Send Emails"):
        cc_email = "vjoenithin@kluniversity.in"  # CC recipient
        for _, row in emails_df.iterrows():
            try:
                msg = MIMEMultipart()
                msg["Subject"] = row["Subject"]
                msg["From"] = outlook_user
                msg["To"] = row["To"]
                msg["CC"] = cc_email  # Adding CC
                msg["X-Priority"] = "1"  # High Priority
                msg["X-MSMail-Priority"] = "High"  # High Priority for Outlook
                msg["Importance"] = "High"  # High Priority for Email Clients
                msg.attach(MIMEText(row["Body"], "html"))  # Use "html" instead of "plain"
                recipients = [row["To"], cc_email]  # Include both main recipient and CC
                if attachment_file:
                    for file in attachment_file:  # Loop through multiple files
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(file.read())  # Read each file
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename={file.name}")
                        msg.attach(part)
                        file.seek(0)  # Reset file pointer

                
                with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
                    server.starttls()
                    server.login(outlook_user, outlook_password)
                    server.sendmail(outlook_user, recipients, msg.as_string())

                st.success(f"📩 High Priority Email sent successfully to {row['To']} (CC: {cc_email})")
            except Exception as e:
                st.error(f"❌ Failed to send email to {row['To']}. Error: {e}")

# Footer
st.markdown("---")
st.markdown("""
    <p style='text-align:center; font-size:14px; color:gray;'>
        Made with ❤️ from <b>Intelligentsia Club</b><br>
        Department of AI & DS<br><br>
        Developed by <b>Aravind</b> (2200080137)<br>
        Contact: <a href='https://t.me/iarvn1' target='_blank'>@iarvn1</a> on Telegram
    </p>
""", unsafe_allow_html=True)
