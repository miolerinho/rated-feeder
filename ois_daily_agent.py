#!/usr/bin/env python3
"""
OIS Daily Agent

Fetches Bank of England OIS rates and sends daily email summaries.
Run this script daily at 7am via cron or a task scheduler.
"""

import sys
import os
from datetime import datetime

# Add python_modules to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'python_modules'))

from boe_ois_fetcher import BOEOISFetcher
from email_sender import EmailSender


def main():
    """Main execution function."""
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting OIS Daily Agent...")

    # Initialize fetcher
    fetcher = BOEOISFetcher()

    # Fetch OIS data
    print("Fetching Bank of England OIS data...")
    data, error = fetcher.fetch_ois_data()

    if error:
        print(f"ERROR: {error}")
        sys.exit(1)

    if not data:
        print("ERROR: No data returned")
        sys.exit(1)

    print("✓ Data fetched successfully")

    # Generate summary
    text_summary = fetcher.format_summary(data)
    print("\n" + text_summary)

    # Initialize email sender
    email_sender = EmailSender()

    # Check if email is configured
    if not email_sender.smtp_user or not email_sender.to_email:
        print("\n⚠ Email not configured - skipping email send")
        print("To enable email, set these environment variables:")
        print("  SMTP_USER, SMTP_PASSWORD, TO_EMAIL")
        print("\nOptional (with defaults):")
        print("  SMTP_HOST (default: smtp.gmail.com)")
        print("  SMTP_PORT (default: 587)")
        print("  FROM_EMAIL (default: same as SMTP_USER)")
        return

    # Send email
    print("\nSending email notification...")

    subject = f"OIS Rates - {data['latest_date'].strftime('%d %B %Y')}"
    html_body = email_sender.create_ois_html_email(data)

    success, error = email_sender.send_email(
        subject=subject,
        body=text_summary,
        html_body=html_body
    )

    if success:
        print(f"✓ Email sent successfully to {email_sender.to_email}")
    else:
        print(f"✗ Email failed: {error}")
        sys.exit(1)

    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] OIS Daily Agent completed")


if __name__ == "__main__":
    main()