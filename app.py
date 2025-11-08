from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import poplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re
from datetime import datetime
from functools import wraps

app = Flask(__name__)
app.secret_key = 'India@team05'  # Change this to a random string

# Email configuration
EMAIL_CONFIG = {
    'pop_server': 'webmailmd.aptaracorp.com',
    'smtp_server': 'webmailmd.aptaracorp.com',
    'pop_port': 110,
    'smtp_port': 587,
    'use_ssl_pop': False,
    'use_tls_smtp': True
}

class EmailManager:
    def __init__(self, email_address, password):
        self.email_address = email_address
        self.password = password
        self.pop_connection = None
        self.smtp_connection = None
    
    def connect_pop(self):
        """Connect to POP3 server"""
        try:
            if EMAIL_CONFIG['use_ssl_pop']:
                self.pop_connection = poplib.POP3_SSL(
                    EMAIL_CONFIG['pop_server'], 
                    EMAIL_CONFIG['pop_port'], 
                    timeout=30
                )
            else:
                self.pop_connection = poplib.POP3(
                    EMAIL_CONFIG['pop_server'], 
                    EMAIL_CONFIG['pop_port'], 
                    timeout=30
                )
            
            self.pop_connection.user(self.email_address)
            self.pop_connection.pass_(self.password)
            return True
        except Exception as e:
            print(f"POP3 connection error: {str(e)}")
            return False
    
    def connect_smtp(self):
        """Connect to SMTP server"""
        try:
            self.smtp_connection = smtplib.SMTP(
                EMAIL_CONFIG['smtp_server'], 
                EMAIL_CONFIG['smtp_port'], 
                timeout=30
            )
            
            if EMAIL_CONFIG['use_tls_smtp']:
                self.smtp_connection.starttls()
            
            self.smtp_connection.login(self.email_address, self.password)
            return True
        except Exception as e:
            print(f"SMTP connection error: {str(e)}")
            return False
    
    def decode_mime_words(self, s):
        """Decode MIME encoded strings"""
        if s is None:
            return ""
        decoded_fragments = decode_header(s)
        fragments = []
        for fragment, encoding in decoded_fragments:
            if isinstance(fragment, bytes):
                try:
                    fragment = fragment.decode(encoding or 'utf-8', errors='ignore')
                except:
                    fragment = fragment.decode('utf-8', errors='ignore')
            fragments.append(str(fragment))
        return ''.join(fragments)
    
    def clean_text(self, text):
        """Clean text for display"""
        if text is None:
            return ""
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def get_email_body(self, msg):
        """Extract email body from message"""
        body = ""
        html_body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        body = str(part.get_payload())
                elif content_type == "text/html" and "attachment" not in content_disposition:
                    try:
                        html_body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        html_body = str(part.get_payload())
        else:
            try:
                content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                if msg.get_content_type() == "text/html":
                    html_body = content
                else:
                    body = content
            except:
                body = str(msg.get_payload())
        
        return body if body else html_body
    
    def get_emails(self, limit=50):
        """Retrieve emails from POP3 server"""
        if not self.connect_pop():
            return None
        
        try:
            num_messages = len(self.pop_connection.list()[1])
            start_msg = max(1, num_messages - limit + 1)
            
            emails = []
            
            for i in range(num_messages, start_msg - 1, -1):
                try:
                    response, lines, octets = self.pop_connection.retr(i)
                    msg_content = b'\r\n'.join(lines)
                    msg = email.message_from_bytes(msg_content)
                    
                    email_data = {
                        'id': i,
                        'date': msg.get('Date', ''),
                        'from': self.decode_mime_words(msg.get('From', '')),
                        'to': self.decode_mime_words(msg.get('To', '')),
                        'subject': self.decode_mime_words(msg.get('Subject', '(No Subject)')),
                        'body': self.get_email_body(msg),
                        'message_id': msg.get('Message-ID', '')
                    }
                    
                    emails.append(email_data)
                except Exception as e:
                    print(f"Error retrieving email {i}: {str(e)}")
                    continue
            
            self.pop_connection.quit()
            return emails
            
        except Exception as e:
            print(f"Error getting emails: {str(e)}")
            return None
    
    def send_email(self, to_address, subject, body, in_reply_to=None):
        """Send an email via SMTP"""
        if not self.connect_smtp():
            return False, "Failed to connect to SMTP server"
        
        try:
            msg = MIMEMultipart()
            msg['From'] = self.email_address
            msg['To'] = to_address
            msg['Subject'] = subject
            
            if in_reply_to:
                msg['In-Reply-To'] = in_reply_to
                msg['References'] = in_reply_to
            
            msg.attach(MIMEText(body, 'plain'))
            
            self.smtp_connection.send_message(msg)
            self.smtp_connection.quit()
            
            return True, "Email sent successfully"
            
        except Exception as e:
            return False, f"Error sending email: {str(e)}"

# Login required decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'email' not in session:
            flash('Please login first', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/')
def index():
    if 'email' in session:
        return redirect(url_for('inbox'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email_address = request.form.get('email')
        password = request.form.get('password')
        
        # Test connection
        manager = EmailManager(email_address, password)
        if manager.connect_pop():
            manager.pop_connection.quit()
            session['email'] = email_address
            session['password'] = password
            flash('Login successful!', 'success')
            return redirect(url_for('inbox'))
        else:
            flash('Login failed. Please check your credentials.', 'danger')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('Logged out successfully', 'info')
    return redirect(url_for('login'))

@app.route('/inbox')
@login_required
def inbox():
    manager = EmailManager(session['email'], session['password'])
    emails = manager.get_emails(limit=50)
    
    if emails is None:
        flash('Error retrieving emails', 'danger')
        return redirect(url_for('logout'))
    
    return render_template('inbox.html', emails=emails)

@app.route('/email/<int:email_id>')
@login_required
def view_email(email_id):
    manager = EmailManager(session['email'], session['password'])
    emails = manager.get_emails(limit=100)
    
    if emails is None:
        flash('Error retrieving email', 'danger')
        return redirect(url_for('inbox'))
    
    # Find the specific email
    email_data = next((e for e in emails if e['id'] == email_id), None)
    
    if email_data is None:
        flash('Email not found', 'warning')
        return redirect(url_for('inbox'))
    
    return render_template('view_email.html', email=email_data)

@app.route('/compose', methods=['GET', 'POST'])
@login_required
def compose():
    if request.method == 'POST':
        to_address = request.form.get('to')
        subject = request.form.get('subject')
        body = request.form.get('body')
        
        manager = EmailManager(session['email'], session['password'])
        success, message = manager.send_email(to_address, subject, body)
        
        if success:
            flash(message, 'success')
            return redirect(url_for('inbox'))
        else:
            flash(message, 'danger')
    
    return render_template('compose.html')

@app.route('/reply/<int:email_id>', methods=['GET', 'POST'])
@login_required
def reply(email_id):
    manager = EmailManager(session['email'], session['password'])
    emails = manager.get_emails(limit=100)
    
    if emails is None:
        flash('Error retrieving email', 'danger')
        return redirect(url_for('inbox'))
    
    # Find the specific email
    original_email = next((e for e in emails if e['id'] == email_id), None)
    
    if original_email is None:
        flash('Email not found', 'warning')
        return redirect(url_for('inbox'))
    
    if request.method == 'POST':
        body = request.form.get('body')
        subject = f"Re: {original_email['subject']}" if not original_email['subject'].startswith('Re:') else original_email['subject']
        
        # Extract email address from "Name <email@domain.com>" format
        from_addr = original_email['from']
        email_match = re.search(r'<(.+?)>', from_addr)
        to_address = email_match.group(1) if email_match else from_addr
        
        success, message = manager.send_email(
            to_address, 
            subject, 
            body, 
            in_reply_to=original_email['message_id']
        )
        
        if success:
            flash(message, 'success')
            return redirect(url_for('inbox'))
        else:
            flash(message, 'danger')
    
    return render_template('reply.html', email=original_email)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)