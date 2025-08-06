import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import openai
import io

st.set_page_config(page_title="Bulk Email Sender", layout="centered")
st.title("📧 Bulk Email Sender (with AI & Attachments)")

# 1. Upload Email List
st.header("1. Upload Email List")
uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    st.success(f"Loaded {len(df)} emails!")
    st.write(df.head())

    # 2. Enter Email Subject
    st.header("2. Email Subject")
    subject = st.text_input("Subject", "Your Marketing Email")

    # 3. Compose Body (AI or Manual)
    st.header("3. Compose Email")
    use_ai = st.checkbox("Use AI to generate email body (OpenAI)", value=False)
    if use_ai:
        openai_api_key = st.text_input("Enter your OpenAI API Key", type="password")
        prompt = st.text_area("Describe your campaign (for AI)", "Announce our new product launch.")
        if st.button("Generate Email Content") and openai_api_key:
            openai.api_key = openai_api_key
            with st.spinner("Generating with AI..."):
                completion = openai.ChatCompletion.create(
                    model="gpt-3.5-turbo",
                    messages=[{"role": "user", "content": prompt}]
                )
                ai_body = completion.choices[0].message.content
            st.text_area("AI-Generated Body", ai_body, height=200)
            body = ai_body
        else:
            body = ""
    else:
        body = st.text_area("Manual Email Body", "Hello,\n\nThis is a test marketing email.")

    # 4. Attach File (Optional)
    st.header("4. Attach a File (Optional)")
    attachment = st.file_uploader("Attach a file (any type)", type=None, key="attach")

    # 5. SMTP Settings
    st.header("5. SMTP Settings")
    smtp_type = st.radio("SMTP Service", ["Outlook", "Brevo (Recommended)"])
    if smtp_type == "Outlook":
        smtp_server = "smtp.office365.com"
        port = 587
        st.info("Outlook: Use your full email. If 2FA, generate App Password.")
    else:
        smtp_server = "smtp-relay.brevo.com"
        port = 587
        st.info("Brevo: Use your Brevo-verified sender and SMTP key.")

    sender_email = st.text_input("Sender Email")
    smtp_password = st.text_input("SMTP Password / Brevo SMTP Key", type="password")

    # 6. Send Emails
    st.header("6. Send Emails")
    if st.button("Send Bulk Emails"):
        if not body or not subject:
            st.error("Please provide both subject and body.")
        else:
            st.write("Sending emails...")
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()
            try:
                server.login(sender_email, smtp_password)
            except Exception as e:
                st.error(f"SMTP login failed: {e}")
                server.quit()
                st.stop()

            success, failed = [], []
            for idx, row in df.iterrows():
                recipient = row['Email']
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'plain'))

                # Add attachment if present
                if attachment is not None:
                    file_bytes = attachment.read()
                    file_part = MIMEApplication(file_bytes, Name=attachment.name)
                    file_part['Content-Disposition'] = f'attachment; filename="{attachment.name}"'
                    msg.attach(file_part)

                try:
                    server.sendmail(sender_email, recipient, msg.as_string())
                    success.append(recipient)
                except Exception as e:
                    failed.append((recipient, str(e)))
            server.quit()

            st.success(f"Sent to {len(success)} emails! Failed for {len(failed)}.")
            st.write("✅ Success:", success[:10], "..." if len(success) > 10 else "")
            if failed:
                st.warning("❌ Failed:")
                st.write(failed[:10])

            # Download summary as CSV
            result_df = pd.DataFrame({'Success': success})
            result_failed = pd.DataFrame(failed, columns=['Failed', 'Reason'])
            st.download_button("Download Success List", result_df.to_csv(index=False), file_name="success.csv")
            st.download_button("Download Failed List", result_failed.to_csv(index=False), file_name="failed.csv")
