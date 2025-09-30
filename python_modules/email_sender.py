"""
Email Sender Module

Sends email notifications with OIS rate summaries.
Supports SMTP configuration for various email providers.
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing import Optional, Tuple
import os


class EmailSender:
    """Handles sending email notifications."""

    def __init__(
        self,
        smtp_host: Optional[str] = None,
        smtp_port: Optional[int] = None,
        smtp_user: Optional[str] = None,
        smtp_password: Optional[str] = None,
        from_email: Optional[str] = None,
        to_email: Optional[str] = None
    ):
        """
        Initialize EmailSender with SMTP configuration.

        Parameters can be provided directly or via environment variables:
        - SMTP_HOST
        - SMTP_PORT
        - SMTP_USER
        - SMTP_PASSWORD
        - FROM_EMAIL
        - TO_EMAIL
        """
        self.smtp_host = smtp_host or os.getenv('SMTP_HOST', 'smtp.gmail.com')
        self.smtp_port = smtp_port or int(os.getenv('SMTP_PORT', '587'))
        self.smtp_user = smtp_user or os.getenv('SMTP_USER', '')
        self.smtp_password = smtp_password or os.getenv('SMTP_PASSWORD', '')
        self.from_email = from_email or os.getenv('FROM_EMAIL', self.smtp_user)
        self.to_email = to_email or os.getenv('TO_EMAIL', '')

    def send_email(
        self,
        subject: str,
        body: str,
        html_body: Optional[str] = None
    ) -> Tuple[bool, Optional[str]]:
        """
        Send an email.

        Args:
            subject: Email subject
            body: Plain text email body
            html_body: Optional HTML version of the email

        Returns:
            Tuple of (success, error_message)
        """
        if not self.smtp_user or not self.smtp_password:
            return False, "SMTP credentials not configured"

        if not self.to_email:
            return False, "Recipient email not configured"

        try:
            # Create message
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = self.from_email
            msg['To'] = self.to_email
            msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S %z')

            # Add plain text part
            text_part = MIMEText(body, 'plain')
            msg.attach(text_part)

            # Add HTML part if provided
            if html_body:
                html_part = MIMEText(html_body, 'html')
                msg.attach(html_part)

            # Send email
            with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
                server.starttls()
                server.login(self.smtp_user, self.smtp_password)
                server.send_message(msg)

            return True, None

        except smtplib.SMTPAuthenticationError:
            return False, "SMTP authentication failed - check username/password"
        except smtplib.SMTPException as e:
            return False, f"SMTP error: {str(e)}"
        except Exception as e:
            return False, f"Error sending email: {str(e)}"

    def create_ois_html_email(self, data: dict) -> str:
        """
        Create an HTML formatted email for OIS rate data.

        Args:
            data: OIS rate data dictionary from BOEOISFetcher

        Returns:
            HTML string
        """
        latest_date = data['latest_date'].strftime('%d %B %Y')
        previous_date = data['previous_date'].strftime('%d %B %Y')

        html = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }}
        .header {{
            background-color: #002147;
            color: white;
            padding: 20px;
            text-align: center;
            border-radius: 5px 5px 0 0;
        }}
        .content {{
            background-color: #f4f4f4;
            padding: 20px;
        }}
        .rate-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background-color: white;
        }}
        .rate-table th {{
            background-color: #002147;
            color: white;
            padding: 12px;
            text-align: left;
        }}
        .rate-table td {{
            padding: 12px;
            border-bottom: 1px solid #ddd;
        }}
        .rate-table tr:hover {{
            background-color: #f5f5f5;
        }}
        .positive {{
            color: #28a745;
            font-weight: bold;
        }}
        .negative {{
            color: #dc3545;
            font-weight: bold;
        }}
        .neutral {{
            color: #6c757d;
            font-weight: bold;
        }}
        .footer {{
            text-align: center;
            padding: 20px;
            font-size: 12px;
            color: #666;
        }}
        .date-info {{
            background-color: white;
            padding: 15px;
            margin: 10px 0;
            border-left: 4px solid #002147;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Bank of England OIS Rates</h1>
        <p>Daily Summary</p>
    </div>
    <div class="content">
        <div class="date-info">
            <strong>Latest Data:</strong> {latest_date}<br>
            <strong>Previous Data:</strong> {previous_date}
        </div>

        <table class="rate-table">
            <thead>
                <tr>
                    <th>Tenor</th>
                    <th>Current Rate</th>
                    <th>Previous Rate</th>
                    <th>Change (bps)</th>
                </tr>
            </thead>
            <tbody>
"""

        for tenor, values in data['rates'].items():
            current = values['current']
            previous = values['previous']
            change = values['change_bps']

            if change is not None:
                if change > 0:
                    change_class = 'positive'
                    arrow = '↑'
                elif change < 0:
                    change_class = 'negative'
                    arrow = '↓'
                else:
                    change_class = 'neutral'
                    arrow = '→'
                change_str = f"{arrow} {abs(change):.2f}"
            else:
                change_class = 'neutral'
                change_str = "N/A"

            html += f"""
                <tr>
                    <td><strong>{tenor}</strong></td>
                    <td>{current:.3f}%</td>
                    <td>{previous:.3f}%</td>
                    <td class="{change_class}">{change_str}</td>
                </tr>
"""

        html += """
            </tbody>
        </table>
    </div>
    <div class="footer">
        <p>Data source: Bank of England</p>
        <p>This is an automated daily summary of OIS rates.</p>
    </div>
</body>
</html>
"""
        return html


def main():
    """Test email sender configuration."""
    sender = EmailSender()

    print("Email Sender Configuration:")
    print(f"SMTP Host: {sender.smtp_host}")
    print(f"SMTP Port: {sender.smtp_port}")
    print(f"SMTP User: {sender.smtp_user or 'Not configured'}")
    print(f"From Email: {sender.from_email or 'Not configured'}")
    print(f"To Email: {sender.to_email or 'Not configured'}")
    print("\nTo configure, set environment variables:")
    print("  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, FROM_EMAIL, TO_EMAIL")


if __name__ == "__main__":
    main()