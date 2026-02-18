"""
Test script for email functionality
Tests email configuration and optionally sends a test email
"""
import sys
import os
import configparser
import smtplib
import datetime
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def get_base_path():
    return Path(__file__).resolve().parent

def test_email_config():
    """Test email configuration loading and validation"""
    print("=" * 60)
    print("EMAIL CONFIGURATION TEST")
    print("=" * 60)
    
    config_path = get_base_path() / "config.ini"
    
    if not config_path.exists():
        print("❌ FAILED: config.ini not found")
        return False
    
    print(f"✓ Config file found: {config_path}")
    
    config = configparser.ConfigParser()
    config.read(config_path)
    
    if 'EMAIL' not in config:
        print("❌ FAILED: No [EMAIL] section in config.ini")
        return False
    
    print("✓ [EMAIL] section found")
    
    email_config = config['EMAIL']
    
    # Check all required settings
    enabled = email_config.get('enabled', 'false').lower() in ['true', '1', 'yes']
    smtp_server = email_config.get('smtp_server', '')
    smtp_port = email_config.get('smtp_port', '587')
    sender_email = email_config.get('sender_email', '')
    sender_password = email_config.get('sender_password', '')
    recipients_str = email_config.get('recipients', '')
    
    print(f"\nConfiguration:")
    print(f"  Enabled: {enabled}")
    print(f"  SMTP Server: {smtp_server}")
    print(f"  SMTP Port: {smtp_port}")
    print(f"  Sender Email: {sender_email}")
    print(f"  Password: {'*' * len(sender_password) if sender_password else '(not set)'}")
    
    recipients = [r.strip() for r in recipients_str.split(',') if r.strip()]
    print(f"  Recipients: {len(recipients)}")
    for i, recipient in enumerate(recipients, 1):
        print(f"    {i}. {recipient}")
    
    # Validation
    print(f"\nValidation:")
    if not enabled:
        print("  ⚠ Email is DISABLED in config")
        return False
    
    print("  ✓ Email is enabled")
    
    if not smtp_server:
        print("  ❌ SMTP server not configured")
        return False
    print(f"  ✓ SMTP server configured")
    
    if not sender_email:
        print("  ❌ Sender email not configured")
        return False
    print(f"  ✓ Sender email configured")
    
    if not sender_password:
        print("  ❌ Sender password not configured")
        return False
    print(f"  ✓ Sender password configured")
    
    if not recipients:
        print("  ❌ No recipients configured")
        return False
    print(f"  ✓ Recipients configured ({len(recipients)})")
    
    print("\n✓ Configuration is valid!")
    return True, email_config, recipients

def test_smtp_connection(email_config):
    """Test SMTP connection without sending email"""
    print("\n" + "=" * 60)
    print("SMTP CONNECTION TEST")
    print("=" * 60)
    
    smtp_server = email_config.get('smtp_server')
    smtp_port = int(email_config.get('smtp_port', '587'))
    sender_email = email_config.get('sender_email')
    sender_password = email_config.get('sender_password')
    
    try:
        print(f"Connecting to {smtp_server}:{smtp_port}...")
        with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
            print("✓ Connected to SMTP server")
            
            print("Starting TLS encryption...")
            server.starttls()
            print("✓ TLS started")
            
            print("Authenticating...")
            server.login(sender_email, sender_password)
            print("✓ Authentication successful")
        
        print("\n✓ SMTP CONNECTION TEST PASSED!")
        return True
        
    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ Authentication failed: {e}")
        print("\nTroubleshooting:")
        print("  - Check your email and password")
        print("  - For Gmail, use an App Password (not your regular password)")
        print("  - Enable 2-factor authentication and generate app password at:")
        print("    https://myaccount.google.com/apppasswords")
        return False
    except smtplib.SMTPException as e:
        print(f"❌ SMTP error: {e}")
        return False
    except Exception as e:
        print(f"❌ Connection error: {e}")
        return False

def send_test_email(email_config, recipients):
    """Send a test email"""
    print("\n" + "=" * 60)
    print("SEND TEST EMAIL")
    print("=" * 60)
    
    response = input("\nDo you want to send a test email? (yes/no): ").strip().lower()
    if response not in ['yes', 'y']:
        print("Test email cancelled")
        return
    
    smtp_server = email_config.get('smtp_server')
    smtp_port = int(email_config.get('smtp_port', '587'))
    sender_email = email_config.get('sender_email')
    sender_password = email_config.get('sender_password')
    
    try:
        # Create test message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = "TEST - PLZ CV Engine Email Configuration"
        
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        body = f"""Hello,

This is a TEST email from PLZ CV Engine.

Email configuration is working correctly!

Test Details:
- Sent from: {sender_email}
- Timestamp: {timestamp}
- SMTP Server: {smtp_server}:{smtp_port}

When a Collection Voucher PDF is generated, it will be automatically
sent to the configured recipients.

Best regards,
PLZ CV Engine Email System
"""
        
        msg.attach(MIMEText(body, 'plain'))
        
        print(f"Sending test email to {len(recipients)} recipient(s)...")
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        print("\n✓ TEST EMAIL SENT SUCCESSFULLY!")
        print(f"\nRecipients:")
        for recipient in recipients:
            print(f"  - {recipient}")
        print("\nCheck your inbox to confirm delivery.")
        
    except Exception as e:
        print(f"❌ Failed to send test email: {e}")

def main():
    """Run all tests"""
    print("\nPLZ CV Engine - Email System Test")
    print(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Test 1: Configuration
    result = test_email_config()
    if not result:
        print("\n❌ Email configuration test failed")
        print("Please check your config.ini file and try again")
        return
    
    _, email_config, recipients = result
    
    # Test 2: SMTP Connection
    if not test_smtp_connection(email_config):
        print("\n❌ SMTP connection test failed")
        print("Email functionality will not work until connection issues are resolved")
        return
    
    # Test 3: Send test email (optional)
    send_test_email(email_config, recipients)
    
    print("\n" + "=" * 60)
    print("ALL TESTS COMPLETED")
    print("=" * 60)
    print("\nEmail functionality is ready to use!")
    print("PDFs will be automatically emailed when generated.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nTest cancelled by user")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
