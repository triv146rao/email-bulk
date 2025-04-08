import dash
import base64
from dotenv import load_dotenv
import os
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import smtplib
from email.message import EmailMessage
import io
import time
from flask import Flask, redirect, request
import uuid

# Initialize the app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# Store uploaded email list and attachments
df = pd.DataFrame()
attachment_data = None
attachment_filename = None
click_logs = []
bounced_emails_cache = []

# Layout
app.layout = dbc.Container([
    html.H1("Email Sender", className="text-center mt-4"),

    dcc.Upload(
        id='upload-data',
        children=html.Button("Upload Excel File"),
        multiple=False,
        className="mb-3"
    ),

    html.Div(id='file-info', className="mb-3"),
    dash_table.DataTable(id='email-table', page_size=5),

    dcc.Input(id='email-subject', type='text', placeholder='Enter email subject', className="form-control mt-3"),
    dcc.Textarea(id='message', placeholder='Enter your message here', className="form-control mt-3"),

    dcc.Upload(
        id='upload-attachment',
        children=html.Button("Upload Attachment"),
        multiple=False,
        className="mb-3"
    ),

    html.Div(id='attachment-info', className="mb-3"),
    dbc.Button("Send Emails", id='send-emails', color='primary', className="mt-3"),
    html.Div(id='status', className="mt-3 text-success"),

    html.H3("Bounced Emails"),
    dash_table.DataTable(id='bounced-emails', page_size=5),
    dbc.Button("Download Bounced Emails", id="download-bounced-btn", color="danger", className="mt-2"),
    dcc.Download(id="download-bounced")
], className="mt-4")

# File upload callback
@app.callback(
    [Output('file-info', 'children'), Output('email-table', 'data'), Output('email-table', 'columns')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_table(contents, filename):
    global df
    if contents is None:
        return "", [], []

    content_type, content_string = contents.split(',')
    decoded = io.BytesIO(base64.b64decode(content_string))
    df = pd.read_excel(decoded)

    if 'Email' not in df.columns:
        return "Error: No 'Email' column found!", [], []

    return f"Uploaded File: {filename}", df.to_dict('records'), [{'name': i, 'id': i} for i in df.columns]

# Attachment upload
@app.callback(
    Output('attachment-info', 'children'),
    Input('upload-attachment', 'contents'),
    State('upload-attachment', 'filename')
)
def store_attachment(contents, filename):
    global attachment_data, attachment_filename
    if contents is None:
        return "No attachment uploaded."

    attachment_filename = filename
    content_type, content_string = contents.split(',')
    attachment_data = base64.b64decode(content_string)
    return f"Uploaded Attachment: {filename}"

# Email sending logic
@app.callback(
    [Output('status', 'children'), Output('bounced-emails', 'data'), Output('bounced-emails', 'columns')],
    Input('send-emails', 'n_clicks'),
    [State('email-subject', 'value'), State('message', 'value')]
)
def send_emails(n_clicks, subject, message):
    global bounced_emails_cache

    if not n_clicks or df.empty:
        return "Please upload an email list and enter a subject & message.", [], []

    if not subject or not message:
        return "Subject and message cannot be empty!", [], []

    load_dotenv()
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    smtp_port = 587
    smtp_server = "smtp.gmail.com"

    bounced_emails = []
    email_batch_size = 50
    delay_between_emails = 2

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)

        for index, row in df.iterrows():
            email = row['Email'].strip().lower()
            msg = EmailMessage()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = subject

            # Create tracking link
            click_id = str(uuid.uuid4())
            track_url = f"http://localhost:8000/track_click/{click_id}?email={email}&redirect=https://example.com"
            tracked_message = f"{message}\n\nClick here: {track_url}"
            msg.set_content(tracked_message)

            if attachment_data and attachment_filename:
                msg.add_attachment(attachment_data, maintype='application', subtype='octet-stream', filename=attachment_filename)

            try:
                response = server.send_message(msg)
                if response:
                    bounced_emails.append({"Email": email, "Error": f"SMTP Response: {response}"})
            except smtplib.SMTPRecipientsRefused:
                bounced_emails.append({"Email": email, "Error": "Recipient refused"})
            except smtplib.SMTPException as e:
                bounced_emails.append({"Email": email, "Error": str(e)})
            except Exception:
                bounced_emails.append({"Email": email, "Error": "Unknown error"})

            time.sleep(delay_between_emails)

            if (index + 1) % email_batch_size == 0:
                server.quit()
                time.sleep(60)
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)

        server.quit()

        # Cache bounced emails for download
        bounced_emails_cache = bounced_emails.copy()

        if bounced_emails:
            bounced_df = pd.DataFrame(bounced_emails)
            bounced_df.to_csv("bounced_emails_log.csv", mode='a', header=False, index=False)

        status_message = "Emails sent successfully!" if not bounced_emails else "Some emails bounced. Check log."
        return status_message, bounced_emails, [{'name': 'Email', 'id': 'Email'}, {'name': 'Error', 'id': 'Error'}]

    except Exception as e:
        return f"Error: {e}", [], []

# Download bounced emails as Excel
@app.callback(
    Output("download-bounced", "data"),
    Input("download-bounced-btn", "n_clicks"),
    prevent_initial_call=True
)
def download_bounced_emails(n_clicks):
    if not bounced_emails_cache:
        return dash.no_update

    output = io.BytesIO()
    bounced_df = pd.DataFrame(bounced_emails_cache)
    bounced_df.to_excel(output, index=False)
    output.seek(0)
    return dcc.send_bytes(output.read(), "bounced_emails.xlsx")

# Link click tracker
@server.route("/track_click/<click_id>")
def track_click(click_id):
    email_clicked = request.args.get("email")
    destination = request.args.get("redirect")
    print(f"Email clicked: {email_clicked}, Link ID: {click_id}")

    with open("click_log.csv", "a") as f:
        f.write(f"{email_clicked},{click_id}\n")

    return redirect(destination)

# Run the app
if __name__ == "__main__":
    app.run_server(debug=False, host="0.0.0.0", port=8000)
