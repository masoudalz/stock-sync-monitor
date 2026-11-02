# Security Policy

## Supported Versions

This project is currently maintained as a personal utility. No formal supported-version schedule is defined.

## Reporting a Vulnerability

Please do not publish credentials, session cookies, private API responses, customer information, or exploitable details in a public issue.

Until a private security contact is configured for this repository, do not submit sensitive vulnerability details publicly. You may open a minimal public issue asking for a private reporting channel.

## Secret Handling

Treat database passwords, API authorization values, CSRF tokens, and browser session cookies as secrets. Do not commit them to source control. Revoke and rotate any credential that may have been exposed, including credentials present in older Git commits.
