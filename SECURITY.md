# Security

## Credentials

Never commit email passwords, app passwords, OAuth tokens, `.env` files, real CVs, or generated reports.

Store credentials in environment variables. The default configuration expects `INTERNSHIP_EMAIL_PASSWORD`.

If a secret is ever committed, revoke it immediately and remove it from Git history before publishing.
