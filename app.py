from flask import Flask, redirect, url_for, session, request, render_template_string
import msal
import requests

app = Flask(__name__)
app.secret_key = 'your_long_randomly_generated_secret_key'  # Replace with your secret key

# Application (client) ID from Azure portal
CLIENT_ID = '85e1f6ad-2e12-416a-a248-ff723760917d'
# Client secret from Azure portal
CLIENT_SECRET = 'fja8Q~U5-TaU~S065p6YKi6Ifbr-QkjaOfE0-b4Q'
# Tenant ID from Azure portal
TENANT_ID = '084a029e-1435-40bc-8201-87ec1b251fb3'
# Authority URL
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
# Scopes required by the app
SCOPE = ['Mail.Read', 'Mail.ReadWrite']
# Redirect URI (must match the one set in the Azure portal)
REDIRECT_URI = 'http://localhost:5000/getAToken'  # Redirect URI path

# Microsoft Graph API endpoint for unread messages
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false'

# MSAL client
msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

@app.route('/')
def index():
    if not session.get('user'):
        return redirect(url_for('login'))
    return redirect(url_for('emails'))

@app.route('/login')
def login():
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)

    if "access_token" in result:
        session['user'] = result.get('id_token_claims')
        session['access_token'] = result['access_token']
        return redirect(url_for('emails'))
    return "Could not authenticate"

@app.route('/emails')
def emails():
    if 'access_token' not in session:
        return redirect(url_for('login'))

    headers = {
        'Authorization': 'Bearer ' + session['access_token']
    }
    response = requests.get(GRAPH_ENDPOINT, headers=headers)
    emails = response.json()

    for email in emails.get('value', []):
        content = email.get('bodyPreview', '')
        attachments = email.get('attachments', [])
        category = categorize_email(content, attachments)
        tag_email(session['access_token'], email['id'], category)

    email_content = ''
    for email in emails.get('value', []):
        email_content += f"<p>From: {email['from']['emailAddress']['name']}<br>"
        email_content += f"Subject: {email['subject']}<br>"
        email_content += f"Received: {email['receivedDateTime']}<br>"
        email_content += f"Body: {email['bodyPreview']}<br>"
        email_content += f"Category: {email['categories']}</p><hr>"

    return render_template_string(f"""
        <h1>Your Emails</h1>
        {email_content}
        <a href="/">Home</a>
    """)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

def categorize_email(content, attachments):
    content_lower = content.lower()

    spam_keywords = ["buy now", "limited sale", "free gift", "urgent action required", "account suspended", "limited time offer"]
    if any(keyword in content_lower for keyword in spam_keywords):
        return 'Spam'

    resume_keywords = ["resume", "cv", "curriculum vitae", "application for position"]
    resume_file_types = ['.pdf', '.doc', '.docx']
    if any(any(keyword in attachment.lower() for keyword in resume_keywords) or
           any(attachment.lower().endswith(file_type) for file_type in resume_file_types)
           for attachment in attachments):
        return 'Resumes'

    application_keywords = ["application", "job application", "cover letter", "submitted for position"]
    if any(keyword in content_lower for keyword in application_keywords):
        return 'Applications'

    hr_keywords = ["benefits", "policy", "human resources", "employee handbook", "payroll", "vacation"]
    if any(keyword in content_lower for keyword in hr_keywords):
        return 'HR'

    meeting_keywords = ["meeting", "schedule", "appointment", "calendar invite", "conference call", "webinar"]
    if any(keyword in content_lower for keyword in meeting_keywords):
        return 'Meetings'

    return 'Uncategorized'

def tag_email(access_token, message_id, category):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }
    data = {
        "categories": [category]
    }
    response = requests.patch(url, headers=headers, json=data)
    if response.status_code == 200:
        print(f"Email {message_id} tagged as {category}")
    else:
        print(f"Failed to tag email {message_id}: {response.status_code} - {response.text}")

if __name__ == "__main__":
    app.run(debug=True)
