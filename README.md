# Microsoft Agent 365 demos

This project combines a Python agent using the Entra ID Auth SDK sidecar, a PowerShell audit of Microsoft Agent 365 identities in an Entra tenant, and archived investment and wealth-management demos.

## Sidecar agent demo

The solution also includes a Python Portfolio Risk Monitor that uses the Microsoft Entra ID Auth SDK sidecar for delegated and autonomous access tokens. The agent does not hold credentials; the sidecar does. This demo is based on Microsoft documentation and has not been run end to end.

### Additional prerequisites

- PowerShell 7 and the Microsoft Graph PowerShell authentication module
- Docker Desktop
- Python 3.10 or later
- The Entra **Agent ID Administrator** role to create the blueprint secret

### Setup and run

1. Create the blueprint and local-development secret:

   ```powershell
   pwsh .\scripts\01-create-blueprint.ps1 -TenantId <tenant-id>
   ```

   This writes the blueprint settings to the ignored `.env` file.
2. In the Entra admin center, create an agent identity from the blueprint. Set its owner and sponsor, then put its client ID in `AGENT_IDENTITY_ID` in `.env`.
3. Register a front-end public client, add the blueprint's delegated `access_agent` scope, grant admin consent, and set its client ID as `FRONTEND_CLIENT_ID` in `.env`.
4. Start the sidecar with `docker compose up` (it binds to localhost only).
5. In another terminal, install dependencies and run the agent:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   python .\agent\demo.py
   ```

The autonomous Graph request may return 403 until the agent identity is granted the required Graph permission. The archived finance demos under `archive/` use their own `archive/requirements.txt`.

## Prerequisites

- PowerShell 5.1 or later

zxz\\
- The Microsoft Graph PowerShell authentication module:

  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```

## Run the audit

From the repository folder, run:

```powershell
.\agent-estate-audit.ps1
```

The script signs in to Microsoft Graph, lists agent identities with their owners and sponsors, and flags identities that have neither. It displays a summary table and writes the full report to `agent-estate-audit.csv` in the current folder. Treat this report as tenant inventory data and do not commit it.

## Archived finance demos

The earlier Python investment and wealth-management demos and their documentation are preserved in [`archive/`](archive/), including the original [`README.md`](archive/README.md) and [`requirements.txt`](archive/requirements.txt).

## License

MIT. See [LICENSE](LICENSE).