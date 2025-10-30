# SharePoint ASPX Crawler

This repository contains a standalone Python script that enumerates `.aspx` pages in a SharePoint Online site.  The crawler authenticates with Microsoft Graph, walks all drives/lists in the site, optionally downloads page content, and exports the findings to CSV.

## Running on a new machine

1. **Install Python** – Python 3.8 or newer is recommended.
2. **Create a virtual environment (optional but recommended):**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```
3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   The optional `pip-system-certs` package allows the script to trust corporate SSL interception certificates automatically.  You can omit it if you do not need that behaviour.
4. **Open `sharepoint.py` and fill in the `USER_SETTINGS` block:** replace the
   placeholder values with your Azure AD tenant/app details and the SharePoint
   site you want to crawl.  Each field is annotated with comments that call out
   whether it is required or optional.
5. **Run the crawler:**
   ```bash
   python sharepoint.py
   ```

### Choosing an authentication mode

* **Device code (default):** keep `AUTH_MODE` set to `"device"` in
  `USER_SETTINGS`.  When you run the script it prompts you with a verification
  URL and code so you can sign in as a user that has access to the SharePoint
  site.  No client secret is needed.
* **Application (client credential) flow:** change `AUTH_MODE` to
  `"application"` and fill in `AZURE_CLIENT_SECRET`.  The script will acquire an
  app-only token using Microsoft Graph.

### Optional runtime flags

Adjust any of these toggles in the `USER_SETTINGS` dictionary if you want to
change the defaults:

* `DOWNLOAD_ASPX`: set to `False` to skip downloading `.aspx` content and only
  record metadata.
* `DOWNLOAD_DIR`: change where downloaded pages are stored (default:
  `./aspx_dump`).
* `EXPORT_CSV`: change the CSV output path (default: `./aspx_index.csv`).
* `SEARCH_FALLBACK`: set to `False` to disable the Microsoft Graph search pass
  when traversal finds zero `.aspx` pages.

### Outputs

By default the crawler writes:

* Downloaded pages under `./aspx_dump/…` (mirroring the drive/list names and relative paths).
* A CSV inventory at `./aspx_index.csv` listing every discovered page.

Adjust these paths with the `--download-dir` and `--export-csv` options if you need to run the tool on a machine where those defaults are not appropriate.

## Troubleshooting

* **Certificate errors on corporate networks:** install `pip-system-certs` (already listed in `requirements.txt`) so the crawler trusts the workstation's corporate root certificates.
* **Authentication failures:** double-check the tenant, client, and site values.  Application mode also requires that the Azure AD app has the correct Microsoft Graph permissions (Sites.Read.All / Files.Read.All) and that admin consent has been granted.

Once dependencies are installed, you can copy this repository—or just `sharepoint/sharepoint.py` and `requirements.txt`—to any machine and follow the same steps above.
