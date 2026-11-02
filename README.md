# Stock Sync Monitor

A Windows desktop application for monitoring inventory differences between a Microsoft SQL Server database and a compatible online store API, then synchronizing supported product stock values.

The application provides a graphical monitoring dashboard, an order/product repository view, API connection settings, browser-assisted session data extraction, and Excel activity logging.

> **Important:** This project is designed for a specific store API and database structure. It is not a universal connector. Before using it with a live store, configure the correct endpoint and credentials, verify that your API supports the expected requests, and test with non-critical products.

## Key Features

- Reads product stock information from Microsoft SQL Server through `pyodbc`.
- Retrieves product and stock information from a compatible store API.
- Compares database stock with online stock and identifies differences.
- Sends supported stock updates to the store API.
- Runs periodic synchronization checks (60-second interval in the current source code).
- Displays synchronization events and status in a desktop GUI.
- Maintains a local repository of tracked products and stock snapshots.
- Creates an Excel activity log using `openpyxl`.
- Provides an API settings window for the API URL, authorization header, CSRF token, and session cookie.
- Includes optional browser-assisted session extraction workflows.

## Application Preview

Screenshots will be added here. If you publish screenshots, make sure they do not reveal store URLs, customer data, product-sensitive information, cookies, API tokens, database names, or other private details.

## Technology Stack

- Python
- Tkinter and CustomTkinter
- Microsoft SQL Server via `pyodbc`
- HTTP requests via `requests`
- Excel logging via `openpyxl`
- Browser cookie utilities via `browser-cookie3`
- Optional browser automation/integration through Playwright and `websocket-client`

## Requirements

- Windows 10 or Windows 11 (the application uses Windows-specific startup-folder behavior).
- Python 3.10 or newer is recommended.
- Microsoft SQL Server access.
- A compatible Microsoft ODBC Driver for SQL Server installed on Windows.
- Access to the target store API and valid credentials.
- Tkinter, normally included with the standard Windows Python installer.

## Installation

1. Install Python from [python.org](https://www.python.org/downloads/). During installation, enable **Add Python to PATH**.
2. Install a Microsoft ODBC Driver for SQL Server. See Microsoft's [ODBC Driver documentation](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server).
3. Download or clone this repository.
4. Open Command Prompt or PowerShell in the project folder.
5. (Recommended) Create and activate a virtual environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

6. Install the Python packages:

   ```powershell
   py -m pip install --upgrade pip
   py -m pip install -r requirements.txt
   ```

## Configuration

### 1. Database connection

The application reads the SQL Server connection string from the `SYNC_DB_CONN` environment variable. If it is not set, the source contains a non-working placeholder connection string.

Set the variable in PowerShell before launching the application. Replace the example values with your own:

```powershell
$env:SYNC_DB_CONN = "DRIVER={ODBC Driver 18 for SQL Server};SERVER=YOUR_SERVER;DATABASE=YOUR_DATABASE;UID=YOUR_USERNAME;PWD=YOUR_PASSWORD;TrustServerCertificate=yes"
```

Use the driver name that is actually installed on your computer. Follow your organization's database security policy; do not commit real connection strings or passwords to Git.

### 2. Store API credentials

Start the application and open **API Settings**. Configure the values required by your store:

- `BASE_URL`: the store's product API endpoint.
- `Authorization`: the API authorization value, if required by the store.
- `X-CSRF-TOKEN`: the CSRF token, if required by the store.
- `Cookie`: the authenticated browser session cookie, if required by the store.

The application saves API URL, cookie, and CSRF settings in a local `api_config.json` file next to the application (with a Desktop fallback in some permission-error cases). **This file can contain active session credentials. Keep it private and never commit it.** The authorization value is configured in the source's `FIXED_HEADERS` dictionary in this version; replace the placeholder locally, or improve the configuration mechanism before distributing the application to other users. Never publish a populated configuration file.

The **automatic browser extraction** feature may require a supported logged-in browser session. Some Chrome workflows expect remote debugging to be available on `127.0.0.1:9222`; the Playwright-based workflow requires the optional Playwright package and a compatible browser installation.

### 3. API compatibility

The application expects product lookup/detail responses and update routes in a particular format. Your store's API must be compatible with the request parameters, response fields, authentication mechanism, and update payloads implemented in `stock_sync.py`. Review and adapt the API functions before connecting a different store platform.

## Run

From the project folder, with the virtual environment activated and database connection configured:

```powershell
py stock_sync.py
```

The application opens a desktop window. The synchronization worker runs in the background and checks for stock differences approximately every 60 seconds while the GUI remains open.

## Local Data and Generated Files

The application creates or uses local files, including:

- `api_config.json` — API URL and session settings; may contain secrets.
- `stock_repo.json` — locally tracked product repository.
- `last_webzi_stock.json` — last retrieved online stock snapshot.
- `last_db_stock.json` — last database stock snapshot.
- `stock_sync_log_*.xlsx` — Excel synchronization logs.

Some data files are stored beside the application and others are stored on the current user's Desktop, according to the paths defined in the source code. These files are excluded by the repository `.gitignore` where applicable. Back them up securely if they are important to your operation.

## Security

- Never commit API tokens, cookies, CSRF tokens, passwords, database connection strings, or populated configuration files.
- Use a dedicated database account with the minimum permissions required.
- Use HTTPS for the store API.
- Treat browser cookies as passwords: anyone who obtains a valid session cookie may be able to access the associated account.
- If credentials were ever committed or shared, revoke/rotate them immediately. Deleting a secret from the latest commit does not remove it from Git history.
- See [SECURITY.md](SECURITY.md) for reporting security issues.

## Limitations and Operational Notes

- The current implementation is Windows-oriented and has not been documented or tested for macOS/Linux.
- SQL Server connectivity depends on a correctly installed ODBC driver and a database schema compatible with the SQL queries in the source.
- API compatibility is store-specific; a successful network connection does not guarantee that product updates will work correctly.
- Browser cookie extraction depends on browser version, session state, permissions, and local browser configuration.
- This repository currently documents the source code; it does not provide a hosted service, installer, or packaged executable.
- Test carefully before enabling synchronization against production inventory.

## Contributing

Bug reports and carefully scoped pull requests are welcome. Please do not include real credentials, customer information, production database exports, or private store data in issues or pull requests. See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

No license has been selected for this repository yet. Unless a license is added, normal copyright restrictions apply and others do not automatically receive permission to reuse, modify, or distribute the code. Choose a license before inviting external reuse.

---

**Project status:** Personal/early-stage utility. Validate database queries, API behavior, and stock-update results in a safe test environment before production use.
