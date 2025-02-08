import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import io
import numpy as np

st.set_page_config(page_title="NSS Email Notification System")
st.title("Automated Email Notification System")

# User inputs for email credentials
outlook_user = st.text_input("Outlook Email Address", "")
outlook_password = st.text_input("Outlook Password", "", type="password")

with st.sidebar:
    st.header("Instructions")
    st.write("""
    - Your file **must contain** the following columns:
      - `event_name`: Name of the event  
      - `venue`: Location of the event  
      - `Time`: Event start time  
      - `event_incharge_name`: Name of the event in-charge  
      - `student_id_number`: Student’s university ID  
    - **Do not include extra columns or empty rows.**
    - The system will generate emails using `student_id_number@kluniversity.in`.
    """)
    # Create a sample dataframe
    sample_data = pd.DataFrame({
        "event_name": ["Tech Talk", "AI Workshop"],
        "venue": ["Indoor Stadium", "Lab 205"],
        "Time": ["1:30 PM", "2:00 PM"],
        "event_incharge_name": ["Aravind", "Rahul"],
        "student_id_number": ["2200080234", "2200080456"]
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

# Upload CSV or Excel file
uploaded_file = st.file_uploader(
    "Upload CSV or Excel file", type=["csv", "xlsx", "xls"]
)

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            data = pd.read_csv(uploaded_file)
        else:
            data = pd.read_excel(uploaded_file, engine="openpyxl")

        # Strip spaces & drop empty rows
        data.columns = data.columns.str.strip()
        data.dropna(how='all', inplace=True)

        st.write("📂 Uploaded Data Preview:", data.head())
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

    emails = []

    for _, row in data.iterrows():
        # Validate required fields
        if pd.isna(row.get("student_id_number")) or pd.isna(row.get("event_name")):
            continue

        student_email = f"{str(row.get('student_id_number')).split('.')[0]}@kluniversity.in"
        if "nan" in student_email:
            continue

        # Define the email template
        email_body = """
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
                    <p>Kindly ensure your presence at the updated venue by <b>[Time]</b> sharp.</p>
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

        emails.append({
            "To": student_email,
            "Subject": f"Venue Update for {row.get('event_name', 'Unknown Event')}",
            "Body": email_body
        })

    emails_df = pd.DataFrame(emails)
    st.write(emails_df)

    # Button to send emails
    if st.button("Send Emails"):
        cc_email = "2200080137@kluniversity.in"
        for _, row in emails_df.iterrows():
            try:
                msg = MIMEMultipart()
                msg["Subject"] = row["Subject"]
                msg["From"] = outlook_user
                msg["To"] = row["To"]
                msg["CC"] = cc_email
                msg.attach(MIMEText(row["Body"], "html"))
                recipients = [row["To"], cc_email]

                with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
                    server.starttls()
                    server.login(outlook_user, outlook_password)
                    server.sendmail(outlook_user, recipients, msg.as_string())

                st.success(f"📩 Email sent successfully to {row['To']} (CC: {cc_email})")
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
