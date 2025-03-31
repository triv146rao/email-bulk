import dash
import base64
import os
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import smtplib
from email.message import EmailMessage
import io
import time
from flask import Flask
import dash

# Initialize the app
#server = Flask(__name__)  # Expose this for Vercel
#app = dash.Dash(__name__, server=server)
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

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
    dash_table.DataTable(id='bounced-emails', page_size=5)
], className="mt-4")

# Store uploaded email list
df = pd.DataFrame()
attachment_data = None
attachment_filename = None

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

@app.callback(
    [Output('status', 'children'), Output('bounced-emails', 'data'), Output('bounced-emails', 'columns')],
    Input('send-emails', 'n_clicks'),
    [State('email-subject', 'value'), State('message', 'value')]
)
def send_emails(n_clicks, subject, message):
    if not n_clicks or df.empty:
        return "Please upload an email list and enter a subject & message.", [], []
    
    if not subject or not message:
        return "Subject and message cannot be empty!", [], []

    sender_email = "triveniaditi@gmail.com"
    sender_password = "dkut nbof tlds hvqs"  # Use an App Password
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    
    bounced_emails = []
    email_batch_size = 50  # Send emails in batches to avoid rate limits
    delay_between_emails = 2  # Delay in seconds
    
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        
        for index, row in df.iterrows():
            email = row['Email'].strip().lower()  # Clean email address
            msg = EmailMessage()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = subject  # Use user-defined subject
            msg.set_content(message)
            
            # Attach file if uploaded
            if attachment_data and attachment_filename:
                msg.add_attachment(attachment_data, maintype='application', subtype='octet-stream', filename=attachment_filename)
            
            try:
                response = server.send_message(msg)
                
                # Log if the SMTP server returns any error response
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
                time.sleep(60)  # Wait 1 minute before starting the next batch
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)
        
        server.quit()
        
        # Log bounced emails
        if bounced_emails:
            bounced_df = pd.DataFrame(bounced_emails)
            bounced_df.to_csv("bounced_emails_log.csv", mode='a', header=False, index=False)  # Append mode
        
        status_message = "Emails sent successfully!" if not bounced_emails else "Some emails bounced. Check log."
        return status_message, bounced_emails, [{'name': 'Email', 'id': 'Email'}, {'name': 'Error', 'id': 'Error'}]
    except Exception as e:
        return f"Error: {e}", [], []

#if __name__ == '__main__':
#    app.run(debug=True)

server = app.server
if __name__ == "__main__":
    app.run_server(debug=False, host="0.0.0.0", port=8000)


