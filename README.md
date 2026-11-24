# Microsoft Agent 365 estate audit

This repository contains a PowerShell script for inventorying Microsoft Agent 365 identities in an Entra tenant and identifying identities without owners or sponsors.

## Prerequisites

- PowerShell 5.1 or later
- A Microsoft Entra account authorized to read agent identities, with the delegated Microsoft Graph permission `AgentIdentity.Read.All` consented
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
