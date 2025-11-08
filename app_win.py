import poplib
import email
from email.header import decode_header
import csv
import re
import getpass

class EmailToCSV:
    def __init__(self, email_address, password, pop_server, port=110, use_ssl=False):
        self.email_address = email_address
        self.password = password
        self.pop_server = pop_server
        self.port = port
        self.use_ssl = use_ssl
        self.mail = None
    
    def connect(self):
        """Connect to the POP3 server"""
        connection_methods = [
            ("POP3 with SSL on port 995", 995, True),
            ("POP3 without SSL on port 110", 110, False),
        ]
        
        for method_name, port, use_ssl in connection_methods:
            try:
                print(f"\nTrying {method_name}...")
                
                if use_ssl:
                    self.mail = poplib.POP3_SSL(self.pop_server, port, timeout=30)
                else:
                    self.mail = poplib.POP3(self.pop_server, port, timeout=30)
                
                print(f"Connected to server, attempting login...")
                self.mail.user(self.email_address)
                self.mail.pass_(self.password)
                
                print(f"✓ Successfully connected using {method_name}")
                return True
                
            except poplib.error_proto as e:
                print(f"✗ Login failed: {str(e)}")
            except Exception as e:
                print(f"✗ Connection failed: {str(e)}")
            
            self.mail = None
        
        print("\n❌ All connection attempts failed.")
        print("\nPossible solutions:")
        print("1. Verify your password is correct")
        print("2. Check if POP3 is enabled in your email settings")
        print("3. Your server might require an app-specific password")
        print("4. Contact IT support for the correct POP3 settings")
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
        """Clean text for CSV output"""
        if text is None:
            return ""
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def get_email_body(self, msg):
        """Extract email body from message"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        body = str(part.get_payload())
                    break
                elif content_type == "text/html" and not body and "attachment" not in content_disposition:
                    try:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        body = str(part.get_payload())
        else:
            try:
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
            except:
                body = str(msg.get_payload())
        
        return self.clean_text(body)
    
    def export_emails(self, output_file='sent_items.csv', limit=None, sent_only=True):
        """Export emails to CSV (POP3 gets all emails, not just sent)"""
        if not self.mail:
            print("Not connected. Please connect first.")
            return
        
        try:
            # Get number of messages
            num_messages = len(self.mail.list()[1])
            print(f"\n✓ Found {num_messages} emails in mailbox")
            
            if sent_only:
                print("⚠ Note: POP3 retrieves all emails in the mailbox.")
                print("   Filtering to show only sent emails (from your address)...")
            
            if limit and limit < num_messages:
                start_msg = num_messages - limit + 1
                print(f"Processing last {limit} emails (from message {start_msg} to {num_messages})")
            else:
                start_msg = 1
                limit = num_messages
                print(f"Processing all {num_messages} emails")
            
            # Prepare CSV file
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['Date', 'From', 'To', 'CC', 'Subject', 'Email_Text']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                processed_count = 0
                sent_count = 0
                
                # Process each email
                for i in range(start_msg, num_messages + 1):
                    try:
                        # Retrieve email
                        response, lines, octets = self.mail.retr(i)
                        
                        # Join all lines to create the email message
                        msg_content = b'\r\n'.join(lines)
                        msg = email.message_from_bytes(msg_content)
                        
                        # Extract email details
                        date = msg.get('Date', '')
                        from_addr = self.decode_mime_words(msg.get('From', ''))
                        to_addr = self.decode_mime_words(msg.get('To', ''))
                        cc_addr = self.decode_mime_words(msg.get('CC', ''))
                        subject = self.decode_mime_words(msg.get('Subject', ''))
                        body = self.get_email_body(msg)
                        
                        # Filter for sent emails if requested
                        if sent_only:
                            # Check if email is from your address
                            if self.email_address.lower() not in from_addr.lower():
                                processed_count += 1
                                continue
                        
                        # Write to CSV
                        writer.writerow({
                            'Date': date,
                            'From': from_addr,
                            'To': to_addr,
                            'CC': cc_addr,
                            'Subject': subject,
                            'Email_Text': body
                        })
                        
                        sent_count += 1
                        processed_count += 1
                        print(f"Processed {processed_count}/{limit}: {subject[:50]}... [SENT]" if sent_only else f"Processed {processed_count}/{limit}: {subject[:50]}...")
                    
                    except Exception as e:
                        print(f"Error processing email {i}: {str(e)}")
                        continue
            
            if sent_only:
                print(f"\n✓ Successfully exported {sent_count} sent emails to {output_file}")
            else:
                print(f"\n✓ Successfully exported {processed_count} emails to {output_file}")
            
        except Exception as e:
            print(f"Error exporting emails: {str(e)}")
    
    def disconnect(self):
        """Disconnect from the server"""
        if self.mail:
            try:
                self.mail.quit()
                print("Disconnected from server")
            except:
                pass


# Main execution
if __name__ == "__main__":
    print("=== Email to CSV Exporter (POP3) ===\n")
    
    # Configuration
    EMAIL = "pawan.sharma1@aptaracorp.com"
    POP_SERVER = "webmailmd.aptaracorp.com"
    
    # Prompt for password securely
    print(f"Email: {EMAIL}")
    print(f"POP3 Server: {POP_SERVER}")
    PASSWORD = getpass.getpass("Enter your password: ")
    
    # Create exporter instance
    exporter = EmailToCSV(EMAIL, PASSWORD, POP_SERVER)
    
    if exporter.connect():
        print("\n=== Connection Successful! ===")
        
        # Ask user preferences
        print("\n⚠ Important: POP3 downloads ALL emails in your mailbox (Inbox).")
        print("   It cannot access the 'Sent Items' folder directly.")
        print("   We will filter to show only emails sent FROM your address.\n")
        
        filter_choice = input("Export only sent emails? (Y/n): ").strip().lower()
        sent_only = filter_choice != 'n'
        
        limit_input = input("How many recent emails to process? (press Enter for all): ").strip()
        limit = int(limit_input) if limit_input else None
        
        # Export emails
        exporter.export_emails(output_file='sent_items.csv', limit=limit, sent_only=sent_only)
        exporter.disconnect()
    else:
        print("\n❌ Could not connect to the server.")
        print("\nPlease verify:")
        print("1. Your password is correct")
        print("2. POP3 is enabled on your email account")
        print("3. The POP3 server address is: webmailmd.aptaracorp.com")
        print("4. You might need an app-specific password")
    
    print("\nDone!")