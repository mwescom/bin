# sharepoint-ldap-sync

A Python command-line tool that:

1. **Downloads** an Excel workbook from SharePoint Online.
2. **Reads** a list of usernames (sAMAccountName) from a configured column.
3. **Queries** Active Directory (LDAP) for each username and retrieves account attributes such as display name, e-mail, department, title, manager, phone number, account status, last logon, and group memberships.
4. **Writes** the retrieved attribute values back into the same workbook (adding header columns automatically).
5. **Uploads** the updated workbook back to SharePoint.

---

## Requirements

| Requirement | Version |
|---|---|
| Python | ≥ 3.10 |
| office365-rest-python-client | 2.5.x |
| ldap3 | 2.9.x |
| openpyxl | 3.1.x |
| python-dotenv | 1.0.x |

Install all dependencies:

```bash
pip install -r requirements.txt
```

---

## Configuration

All settings are read from environment variables.  Copy `.env.example` to
`.env` and fill in the values:

```bash
cp .env.example .env
$EDITOR .env
```

### SharePoint settings

| Variable | Required | Description |
|---|---|---|
| `SP_SITE_URL` | ✅ | Full URL of the SharePoint site, e.g. `https://contoso.sharepoint.com/sites/MySite` |
| `SP_CLIENT_ID` | ✅ | Azure AD application (client) ID |
| `SP_CLIENT_SECRET` | ✅ | Azure AD application secret |
| `SP_FILE_PATH` | ✅ | Path to the Excel file inside the document library, e.g. `Shared Documents/HR/user_list.xlsx` |

#### Azure AD app registration

The app needs the **Sites.ReadWrite.All** (or **Files.ReadWrite.All**) application
permission granted in Azure Active Directory, and the tenant admin must grant
admin consent.

### LDAP / Active Directory settings

| Variable | Required | Default | Description |
|---|---|---|---|
| `LDAP_SERVER` | ✅ | — | DC hostname or IP |
| `LDAP_PORT` | | `389` | LDAP port (use `636` with `LDAP_USE_SSL=true`) |
| `LDAP_USE_SSL` | | `false` | `true` to enable LDAPS |
| `LDAP_BIND_DN` | ✅ | — | Service account – `DOMAIN\user` for NTLM or full DN for SIMPLE bind |
| `LDAP_BIND_PASSWORD` | ✅ | — | Service account password |
| `LDAP_SEARCH_BASE` | ✅ | — | Base DN for searches, e.g. `DC=contoso,DC=com` |
| `LDAP_ATTRIBUTES` | | (see below) | Comma-separated list of attributes to retrieve |

Default attributes when `LDAP_ATTRIBUTES` is not set:

```
displayName, mail, department, title, manager, telephoneNumber,
userAccountControl, accountExpires, whenCreated, lastLogonTimestamp, memberOf
```

### Excel layout settings

| Variable | Required | Default | Description |
|---|---|---|---|
| `EXCEL_USERNAME_COLUMN` | | `A` | Column letter containing the usernames |
| `EXCEL_HEADER_ROW` | | `1` | Row number of the header row |
| `EXCEL_DATA_START_ROW` | | `2` | First row that contains a username |

---

## Excel workbook format

The tool expects (and produces) a workbook with a structure like:

| A (Username) | displayName | mail | department | … |
|---|---|---|---|---|
| jsmith | John Smith | jsmith@contoso.com | Engineering | … |
| mjones | Mary Jones | mjones@contoso.com | HR | … |

* The **Username** column is configured by `EXCEL_USERNAME_COLUMN` (default: `A`).
* All other columns are created automatically on the first run and reused on
  subsequent runs.
* A **Status** column is added: `OK` when the user was found, `NOT FOUND`
  otherwise.

---

## Usage

```bash
# Run the sync
python main.py
```

The script exits with code `0` on success and `1` on any fatal error.

### Example output

```
2024-03-15 09:00:01 [INFO] main: Step 1/4 – Downloading workbook from SharePoint …
2024-03-15 09:00:03 [INFO] sharepoint_client: Connected to SharePoint site: HR Portal
2024-03-15 09:00:04 [INFO] sharepoint_client: Downloaded 24576 bytes.
2024-03-15 09:00:04 [INFO] main: Step 2/4 – Parsing usernames from workbook …
2024-03-15 09:00:04 [INFO] excel_handler: Workbook loaded; active sheet: Sheet
2024-03-15 09:00:04 [INFO] main: Found 3 username(s) to process.
2024-03-15 09:00:04 [INFO] main: Step 3/4 – Querying LDAP and updating workbook …
2024-03-15 09:00:04 [INFO] ldap_client: Connected to LDAP server dc01.contoso.com:389 (SSL=False)
2024-03-15 09:00:04 [INFO] main:   Processing user: jsmith (row 2)
2024-03-15 09:00:04 [INFO] main:   Processing user: mjones (row 3)
2024-03-15 09:00:04 [WARNING] ldap_client: User 'baduser' not found in LDAP.
2024-03-15 09:00:04 [INFO] main:   Processing user: baduser (row 4)
2024-03-15 09:00:05 [INFO] main: Updated: 2  |  Not found: 1  |  Total: 3
2024-03-15 09:00:05 [INFO] main: Step 4/4 – Uploading updated workbook to SharePoint …
2024-03-15 09:00:06 [INFO] sharepoint_client: Upload complete.
2024-03-15 09:00:06 [INFO] main: Sync complete.
```

---

## Running tests

```bash
pip install pytest
pytest tests/ -v
```

---

## Project structure

```
sharepoint-ldap-sync/
├── main.py               # Entry point / workflow orchestration
├── config.py             # Configuration loading from environment variables
├── sharepoint_client.py  # SharePoint file download / upload
├── ldap_client.py        # LDAP / Active Directory user queries
├── excel_handler.py      # Excel workbook read / write
├── requirements.txt      # Python dependencies
├── .env.example          # Configuration template
└── tests/
    ├── test_config.py
    ├── test_excel_handler.py
    ├── test_ldap_client.py
    └── test_sharepoint_client.py
```
