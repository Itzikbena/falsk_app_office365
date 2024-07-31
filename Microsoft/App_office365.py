import os
import json
import base64
import asyncio
import re
from flask import Flask, redirect, url_for, session, request, jsonify
from flask_session import Session
import msal
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from apscheduler.schedulers.background import BackgroundScheduler
from aiohttp import ClientSession
from datetime import datetime, timezone
import hashlib
import ftplib
import logging
from flasgger import Swagger

# Set up the logger
logging.basicConfig(filename='app.log', level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger(__name__)

# Suppress only the single InsecureRequestWarning from urllib3
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = 'a_random_secret_key'

# Initialize Swagger
swagger = Swagger(app)

# Ensure the directories for token and attachment storage exist
os.makedirs('tokens', exist_ok=True)
os.makedirs('attachments', exist_ok=True)
os.makedirs('last_checked', exist_ok=True)

# Configuration for Flask-Session
app.config['SESSION_FILE_DIR'] = 'flask_session'
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_PERMANENT'] = False

Session(app)

# Configuration for OAuth
CLIENT_ID = '958a0438-82d4-42ae-a9ad-04eca03a939b'
CLIENT_SECRET = 'Xjh8Q~sxjms-ReyJZssrtei.1FFgmUoTtoWIPdw4'
AUTHORITY = 'https://login.microsoftonline.com/common'
REDIRECT_URI = 'https://127.0.0.1:5000/getAToken'
SCOPE = ['Mail.Read']

scheduler = BackgroundScheduler()
scheduler.start()

keywords = ['invoice', 'payment', 'receipt', 'bill', 'statement', 'purchase', 'order', 'transaction', 'confirmation', 'paid']


def upload_to_ftp(file_path, file_name):
    ftp_server = 'ftpn2.deltahost.com.ua'
    ftp_user = 'basz1'
    ftp_password = '25jAXhcdkdR#'
    upload_path = '/desired/upload/path'  # Replace with your desired upload path

    with ftplib.FTP(ftp_server, ftp_user, ftp_password) as ftp:
        # Ensure the directory exists or create it
        try:
            ftp.cwd(upload_path)
        except ftplib.error_perm:
            # Directory does not exist, create it
            parts = upload_path.split('/')
            for part in parts:
                if part:  # Avoid empty parts from leading /
                    try:
                        ftp.cwd(part)
                    except ftplib.error_perm:
                        ftp.mkd(part)
                        ftp.cwd(part)

        # Change to the upload directory
        ftp.cwd(upload_path)

        # Upload the file
        with open(file_path, 'rb') as file:
            ftp.storbinary(f'STOR {file_name}', file)

    return f'ftp://{ftp_server}{upload_path}/{file_name}'

def save_token_to_file(token_cache, user_name):
    token_file = os.path.join('tokens', f'{user_name}.json')
    with open(token_file, 'w') as f:
        f.write(token_cache.serialize())

def load_token_from_file(user_name):
    token_file = os.path.join('tokens', f'{user_name}.json')
    if os.path.exists(token_file):
        with open(token_file, 'r') as f:
            return f.read()
    return None

def save_last_checked(user_name, timestamp):
    last_checked_file = os.path.join('last_checked', f'{user_name}.json')
    with open(last_checked_file, 'w') as f:
        f.write(json.dumps({"last_checked": timestamp}))

def load_last_checked(user_name):
    last_checked_file = os.path.join('last_checked', f'{user_name}.json')
    if os.path.exists(last_checked_file):
        with open(last_checked_file, 'r') as f:
            return json.load(f).get('last_checked')
    return None

async def process_emails(session, user_name):
    logger.info(f"Checking emails for: {user_name}")
    token_cache = load_token_from_file(user_name)
    if not token_cache:
        logger.info(f"No token found for {user_name}, user might need to log in again.")
        return

    cache = msal.SerializableTokenCache()
    cache.deserialize(token_cache)
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if not accounts:
        logger.info(f"No accounts found in cache for {user_name}.")
        return

    token = cca.acquire_token_silent(SCOPE, account=accounts[0])
    if not token:
        logger.info(f"Could not acquire token silently for {user_name}.")
        return

    last_checked = load_last_checked(user_name)
    if last_checked:
        last_checked_dt = datetime.fromisoformat(last_checked)
        params = {
            '$top': '10',
            '$select': 'receivedDateTime,subject,from,body,hasAttachments',
            '$filter': f"receivedDateTime gt {last_checked_dt.isoformat()}"
        }
    else:
        params = {
            '$top': '10',
            '$select': 'receivedDateTime,subject,from,body,hasAttachments'
        }

    headers = {'Authorization': 'Bearer ' + token['access_token']}
    async with session.get('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages',
                           headers=headers, params=params, ssl=False) as resp:
        if resp.status == 200:
            emails = await resp.json()
            await handle_emails(emails.get('value', []), user_name, headers)
        else:
            logger.error(f"Failed to fetch emails for {user_name}: {await resp.text()}")

    save_last_checked(user_name, datetime.now(timezone.utc).isoformat())

async def handle_emails(emails, user_name, headers):
    found_email = False  # Track if any email meets the requirement
    for email in emails:
        for keyword in keywords:
            if keyword in email['subject'].lower() or keyword in email['body']['content'].lower():
                logger.info(f"New email found that meets the requirement for {user_name} with keyword '{keyword}': {email['subject']}")
                found_email = True

        # Check for attachments
        if email['hasAttachments']:
            attachments_checked = await handle_attachments(email, user_name, headers)
            if attachments_checked:
                found_email = True  # Reset if no matching attachments found

                # Check for GreenInvoice link in the body
        greeninvoice_link = find_greeninvoice_link(email['body']['content'])
        if greeninvoice_link:
            await download_greeninvoice_pdf(greeninvoice_link, user_name, email)
            found_email = True

            # Log email details
            with open(os.path.join('attachments', f'{user_name}_emails.txt'), 'a') as f:
                f.write(f"From: {email['from']['emailAddress']['address']} ")
                f.write(f"To: {user_name}\n")

                # Break the loop after finding the first matching keyword
                break
    if not found_email:
        logger.info(f"No new email for {user_name}")

def find_greeninvoice_link(content):
    """Find the GreenInvoice link in the email content."""
    match = re.search(r'https://www\.greeninvoice\.co\.il/api/v1/documents/[^\s]+', content)
    return match.group(0) if match else None


async def download_greeninvoice_pdf(url, user_name, email):
    async with ClientSession() as session:
        async with session.get(url, ssl=False) as resp:  # Disabled SSL verification for testing
            if resp.status == 200:
                content_bytes = await resp.read()
                # Use a hash of the URL to create a unique but shorter file name
                url_hash = hashlib.md5(url.encode()).hexdigest()
                file_name = f"{url_hash}.pdf"
                file_path = os.path.join('attachments', user_name, file_name)
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                with open(file_path, 'wb') as f:
                    f.write(content_bytes)

                # Upload to FTP
                ftp_url = upload_to_ftp(file_path, file_name)

                # Create JSON file with URL and email details
                json_data = {
                    "url": ftp_url,
                    "from": email['from']['emailAddress']['address'],
                    "to": user_name
                }
                json_file_path = os.path.join('attachments', user_name, f"{url_hash}_details.json")
                with open(json_file_path, 'w') as json_file:
                    json.dump(json_data, json_file)

                logger.info(f"Downloaded and uploaded GreenInvoice PDF for {user_name}: {file_name}")  # Added print statement
            else:
                logger.info(f"Failed to download PDF from {url} for {user_name}: {resp.status}")  # Added print statement


async def handle_attachments(email, user_name, headers):
    async with ClientSession() as session:
        email_id = email['id']
        url = f'https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments'
        async with session.get(url, headers=headers, ssl=False) as resp:
            if resp.status == 200:
                attachments = await resp.json()
                found_attachment = False
                for attachment in attachments.get('value', []):
                    if attachment['name'].endswith('.pdf') and any(keyword in attachment['name'].lower() for keyword in keywords):
                        logger.info(f"Attachment {attachment['name']} matches the requirement for {user_name}.")  # Added print statement
                        found_attachment = True
                        await save_attachment(attachment, user_name, email)
                if not found_attachment:
                    logger.info(f"No matching attachments found for {user_name}.")  # Added print statement
                return found_attachment


async def save_attachment(attachment, user_name, email):
    content_bytes = base64.b64decode(attachment['contentBytes'])
    file_name = attachment['name']
    file_path = os.path.join('attachments', user_name, file_name)
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    with open(file_path, 'wb') as f:
        f.write(content_bytes)

    # Upload to FTP
    ftp_url = upload_to_ftp(file_path, file_name)

    # Create JSON file with URL and email details
    json_data = {
        "url": ftp_url,
        "from": email['from']['emailAddress']['address'],
        "to": user_name
    }
    json_file_path = os.path.join('attachments', user_name, f"{os.path.splitext(file_name)[0]}_details.json")
    with open(json_file_path, 'w') as json_file:
        json.dump(json_data, json_file)

    logger.info(f"Saved and uploaded attachment for {user_name}: {file_name}")

async def check_new_emails():
    async with ClientSession() as session:
        users = [user for user in os.listdir('tokens') if user.endswith('.json')]
        tasks = [process_emails(session, user.replace('.json', '')) for user in users]
        await asyncio.gather(*tasks)

@app.route('/log', methods=['GET'])
def get_log():
    """
    Get application log
    ---
    tags:
      - Log
    responses:
      200:
        description: Application log
        content:
          application/json:
            schema:
              type: object
              properties:
                log:
                  type: string
      404:
        description: Log file not found
    """
    log_file_path = 'app.log'
    if os.path.exists(log_file_path):
        with open(log_file_path, 'r') as log_file:
            log_content = log_file.read()
        return jsonify({"log": log_content}), 200
    else:
        return jsonify({"error": "Log file not found"}), 404

@app.route('/delete', methods=['POST'])
def delete_client():
    """
    Delete client data
    ---
    tags:
      - Client
    parameters:
      - name: email
        in: body
        type: string
        required: true
        description: The email of the client to delete
        schema:
          type: object
          required:
            - email
          properties:
            email:
              type: string
    responses:
      200:
        description: Client data and job deleted
      400:
        description: Email not provided
      404:
        description: Email not found
    """
    data = request.json
    email = data.get('email')
    if not email:
        return jsonify({"error": "Email not provided"}), 400

    # Check if email exists in tokens and last_checked folders
    token_file = os.path.join('tokens', f'{email}.json')
    last_checked_file = os.path.join('last_checked', f'{email}.json')

    if not os.path.exists(token_file) and not os.path.exists(last_checked_file):
        return jsonify({"error": "Email not found"}), 404

    # Delete token file if it exists
    if os.path.exists(token_file):
        os.remove(token_file)

    # Delete last_checked file if it exists
    if os.path.exists(last_checked_file):
        os.remove(last_checked_file)

    # Stop the job for this client
    job_id = f'check_emails_{email}'
    if scheduler.get_job(job_id):
        scheduler.remove_job(job_id)

    return jsonify({"message": f"Deleted data and stopped job for {email}"}), 200


@app.route('/<path:url_return>')
def dynamic_route(url_return):
    status_code = request.args.get('status_code', default=200, type=int)

    if status_code == 200:
        logger.info(f"Redirected to: {url_return}")
        return "Redirect successful", status_code
    else:
        logger.error(f"Failed to redirect to: {url_return}")
        return "Redirect unsuccessful", status_code

@app.route('/office365/<url_return>')
def login(url_return):
    session['url_return'] = url_return
    session['flow'] = _build_auth_code_flow(scopes=SCOPE)
    return redirect(session['flow']['auth_uri'])

@app.route('/getAToken')
def authorized():
    try:
        cache = msal.SerializableTokenCache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(session.get('flow', {}), request.args)
        if 'error' in result:
            return "Error: " + result['error']
        session['user'] = result.get('id_token_claims', {})
        user_name = session['user'].get('name', session['user'].get('preferred_username', 'unknown_user'))
        save_token_to_file(cache, user_name)
        # Save current time as last checked
        save_last_checked(user_name, datetime.now(timezone.utc).isoformat())
        job_id = f'check_emails_{user_name}'
        scheduler.add_job(lambda: asyncio.run(check_new_emails()), 'interval', minutes=1, id=job_id)  # Modified
    except ValueError:  # Usually caused by CSRF
        pass  # Ignore for now
    return redirect(url_for('mail'))

@app.route('/list', methods=['GET'])
def list_clients():
    token_files = [f.replace('.json', '') for f in os.listdir('tokens') if f.endswith('.json')]
    return jsonify(token_files), 200



@app.route('/mail')
def mail():
    user_name = session['user'].get('name', session['user'].get('preferred_username', 'unknown_user'))
    token_cache = load_token_from_file(user_name)
    if not token_cache:
        return jsonify({"error": "User not authenticated"}), 401

    cache = msal.SerializableTokenCache()
    cache.deserialize(token_cache)
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if not accounts:
        return jsonify({"error": "No accounts found in cache"}), 402

    token = cca.acquire_token_silent(SCOPE, account=accounts[0])
    if not token:
        return jsonify({"error": "Could not acquire token silently"}), 403

    response = requests.get(
        'https://graph.microsoft.com/v1.0/me/messages',
        headers={'Authorization': 'Bearer ' + token['access_token']},
        params={
            '$top': '10',
            '$select': 'receivedDateTime,subject,from'
        },
        verify=False
    )

    url_return = session.get('url_return', '/')

    # Check if email exists in tokens and last_checked folders
    token_file = os.path.join('tokens', f'{user_name}.json')
    last_checked_file = os.path.join('last_checked', f'{user_name}.json')

    if response.status_code == 200:
        return redirect(url_for('dynamic_route', url_return=url_return, status_code=200))
    else:
        # Delete token file if it exists
        if os.path.exists(token_file):
            os.remove(token_file)

        # Delete last_checked file if it exists
        if os.path.exists(last_checked_file):
            os.remove(last_checked_file)

        # Stop the job for this client
        job_id = f'check_emails_{user_name}'
        if scheduler.get_job(job_id):
            scheduler.remove_job(job_id)
        return redirect(url_for('dynamic_route', url_return=url_return,status_code=500))

def _build_msal_app(cache=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET, token_cache=cache)

def _build_auth_code_flow(scopes=None):
    return _build_msal_app().initiate_auth_code_flow(
        scopes or [],
        redirect_uri=REDIRECT_URI)

if __name__ == '__main__':
    scheduler.add_job(lambda: asyncio.run(check_new_emails()), 'interval', minutes=1, id='check_emails')  # Added
    app.run(ssl_context=('cert.pem', 'key.pem'))
