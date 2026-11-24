#!/usr/bin/env python3
"""
agent365_governance_demo.py
Microsoft Agent 365

Entra Agent ID governance queries via Microsoft Graph.
Extends agent365_finance_demo.py with:
  - Owner and sponsor assignment queries
  - Access package status checks
  - Ownerless agent detection across the tenant
  - Lifecycle state reporting

Required permissions (app registration):
  - Application.Read.All
  - AppRoleAssignment.ReadWrite.All
  - EntitlementManagement.Read.All
  - Directory.Read.All

Setup:
  pip install msal requests rich tabulate

Usage:
  python agent365_governance_demo.py \
    --tenant-id <your-tenant-id> \
    --client-id <your-app-client-id> \
    --client-secret <your-client-secret>
"""

import argparse
import json
import sys
import time
from datetime import datetime, timedelta

import requests
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich import box
from msal import ConfidentialClientApplication

console = Console()

# Finance agents Repo
FINANCE_AGENTS = [
    "WealthAdvisorCopilot",
    "PortfolioRiskMonitor",
    "ComplianceGuardian",
]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"


# ─────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────

def get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Acquire an OAuth 2.0 token via MSAL client credentials flow."""
    app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        error = result.get("error_description", "Unknown auth error")
        console.print(f"[red]❌ Auth failed: {error}[/red]")
        sys.exit(1)
    return result["access_token"]


def graph_get(token: str, url: str, use_beta: bool = False) -> dict:
    """Make an authenticated GET request to Microsoft Graph."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.get(url, headers=headers, timeout=30)
    if response.status_code == 404:
        return {}
    if response.status_code not in (200, 201):
        console.print(f"[yellow]  ⚠ Graph call returned {response.status_code}: {url}[/yellow]")
        return {}
    return response.json()


# ─────────────────────────────────────────────
# Module 1: Agent app registration lookup
# ─────────────────────────────────────────────

def find_agent_app(token: str, agent_name: str) -> dict | None:
    """Find an app registration by display name."""
    url = f"{GRAPH_BASE}/applications?$filter=displayName eq '{agent_name}'&$select=id,appId,displayName,createdDateTime,signInAudience,requiredResourceAccess"
    data = graph_get(token, url)
    apps = data.get("value", [])
    return apps[0] if apps else None


def get_service_principal(token: str, app_id: str) -> dict | None:
    """Get the service principal (enterprise app) for an app registration."""
    url = f"{GRAPH_BASE}/servicePrincipals?$filter=appId eq '{app_id}'&$select=id,displayName,accountEnabled,createdDateTime,appRoleAssignments"
    data = graph_get(token, url)
    sps = data.get("value", [])
    return sps[0] if sps else None


def get_app_owners(token: str, app_object_id: str) -> list[dict]:
    """Get owners assigned to an app registration."""
    url = f"{GRAPH_BASE}/applications/{app_object_id}/owners?$select=id,displayName,userPrincipalName,mail"
    data = graph_get(token, url)
    return data.get("value", [])


# ─────────────────────────────────────────────
# Module 2: Access packages (Entitlement Management)
# ─────────────────────────────────────────────

def get_access_packages(token: str, agent_name: str) -> list[dict]:
    """Find access packages that contain the agent name."""
    url = f"{GRAPH_BASE}/identityGovernance/entitlementManagement/accessPackages?$select=id,displayName,description,createdDateTime,isHidden"
    data = graph_get(token, url)
    packages = data.get("value", [])
    # Filter to those related to this agent
    return [p for p in packages if agent_name.lower() in p.get("displayName", "").lower()]


def get_access_package_assignments(token: str, package_id: str) -> list[dict]:
    """Get current assignments for an access package."""
    url = f"{GRAPH_BASE}/identityGovernance/entitlementManagement/assignments?$filter=accessPackage/id eq '{package_id}'&$expand=target&$select=id,state,expiredDateTime,target"
    data = graph_get(token, url)
    return data.get("value", [])


# ─────────────────────────────────────────────
# Module 3: Ownerless agent detection
# ─────────────────────────────────────────────

def detect_ownerless_agents(token: str, agent_names: list[str]) -> list[dict]:
    """
    Detect agents that have no owner assigned.
    An ownerless agent is a governance gap — no human is accountable
    for its lifecycle or access decisions.
    """
    ownerless = []
    for name in agent_names:
        app = find_agent_app(token, name)
        if not app:
            continue
        owners = get_app_owners(token, app["id"])
        if not owners:
            ownerless.append({
                "agent": name,
                "app_id": app.get("appId"),
                "created": app.get("createdDateTime"),
                "risk": "HIGH — No owner assigned",
            })
    return ownerless


# ─────────────────────────────────────────────
# Module 4: Lifecycle state report
# ─────────────────────────────────────────────

def get_sp_signin_activity(token: str, sp_id: str) -> dict | None:
    """
    Get last sign-in activity for a service principal.
    Uses beta endpoint — signInActivity not available in v1.0.
    """
    url = f"{GRAPH_BETA}/servicePrincipals/{sp_id}?$select=id,displayName,signInActivity"
    data = graph_get(token, url, use_beta=True)
    return data.get("signInActivity")


# ─────────────────────────────────────────────
# Report rendering
# ─────────────────────────────────────────────

def render_agent_identity_table(agents_data: list[dict]):
    """Render a Rich table showing identity governance state per agent."""
    table = Table(
        title="Agent Identity Governance Report",
        box=box.ROUNDED,
        border_style="blue",
        header_style="bold cyan",
        show_lines=True,
    )
    table.add_column("Agent", style="bold white", no_wrap=True)
    table.add_column("App ID", style="dim", max_width=36)
    table.add_column("SP Enabled", justify="center")
    table.add_column("Owners", justify="center")
    table.add_column("Access Packages", justify="center")
    table.add_column("Last Sign-In", max_width=20)
    table.add_column("Governance Status")

    for d in agents_data:
        sp_enabled = "✅" if d.get("sp_enabled") else "❌"
        owner_count = d.get("owner_count", 0)
        owner_cell = f"✅ {owner_count}" if owner_count > 0 else "[red]❌ None[/red]"
        pkg_count = d.get("access_package_count", 0)
        pkg_cell = f"✅ {pkg_count}" if pkg_count > 0 else "[yellow]⚠ None[/yellow]"
        last_signin = d.get("last_signin", "No activity")

        # Derive governance status
        issues = []
        if owner_count == 0:
            issues.append("[red]No Owner[/red]")
        if pkg_count == 0:
            issues.append("[yellow]No Access Package[/yellow]")
        if not d.get("sp_enabled"):
            issues.append("[dim]SP Disabled[/dim]")
        status = " | ".join(issues) if issues else "[green]Governed[/green]"

        table.add_row(
            d["name"],
            d.get("app_id", "Not found"),
            sp_enabled if d.get("app_id") else "—",
            owner_cell if d.get("app_id") else "—",
            pkg_cell if d.get("app_id") else "—",
            last_signin,
            status if d.get("app_id") else "[dim]Not registered[/dim]",
        )

    console.print(table)


def render_ownerless_report(ownerless: list[dict]):
    """Render ownerless agent findings."""
    if not ownerless:
        console.print(Panel("✅ No ownerless agents detected.", style="green"))
        return

    table = Table(
        title=f"⚠  Ownerless Agents Detected ({len(ownerless)})",
        box=box.ROUNDED,
        border_style="red",
        header_style="bold red",
    )
    table.add_column("Agent")
    table.add_column("App ID")
    table.add_column("Created")
    table.add_column("Risk")

    for item in ownerless:
        created = item.get("created", "Unknown")
        if created and created != "Unknown":
            created = created[:10]
        table.add_row(
            item["agent"],
            item.get("app_id", ""),
            created,
            f"[red]{item['risk']}[/red]",
        )

    console.print(table)
    console.print(
        "\n[yellow]Remediation:[/yellow] Assign an Owner in Entra admin center → "
        "App registrations → [agent] → Owners. "
        "Then create an Access Package with Sponsor approval via "
        "Entra admin center → Identity Governance → Entitlement Management.\n"
    )


def render_access_package_report(pkgs_data: list[dict]):
    """Render access package assignment details."""
    if not pkgs_data:
        console.print("[yellow]  No access packages found for finance agents.[/yellow]")
        return

    table = Table(
        title="Access Package Assignments",
        box=box.ROUNDED,
        border_style="cyan",
        header_style="bold cyan",
    )
    table.add_column("Package")
    table.add_column("State")
    table.add_column("Expires")
    table.add_column("Assignee")

    for pkg in pkgs_data:
        for assignment in pkg.get("assignments", []):
            target = assignment.get("target", {})
            assignee = target.get("displayName", target.get("email", "Unknown"))
            expires = assignment.get("expiredDateTime", "No expiry set")
            if expires and expires != "No expiry set":
                expires = expires[:10]
            state = assignment.get("state", "unknown")
            state_cell = f"[green]{state}[/green]" if state == "delivered" else f"[yellow]{state}[/yellow]"

            table.add_row(pkg["displayName"], state_cell, expires, assignee)

    console.print(table)


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Agent 365 — Entra Identity Governance Demo"
    )
    parser.add_argument("--tenant-id", required=True, help="Azure AD tenant ID")
    parser.add_argument("--client-id", required=True, help="App registration client ID")
    parser.add_argument("--client-secret", required=True, help="App registration client secret")
    parser.add_argument(
        "--agents",
        nargs="+",
        default=FINANCE_AGENTS,
        help="Agent display names to audit (default: finance agents repo)",
    )
    args = parser.parse_args()

    console.print(Panel.fit(
        "[bold cyan]Microsoft Agent 365[/bold cyan]\n"
        "[white]Entra Agent ID Governance Demo[/white]\n"
        f"[dim]Auditing {len(args.agents)} agents | {datetime.now().strftime('%Y-%m-%d %H:%M')}[/dim]",
        border_style="cyan",
    ))

    # Authenticate
    console.print("\n[bold]Step 1:[/bold] Authenticating with Microsoft Graph...")
    token = get_token(args.tenant_id, args.client_id, args.client_secret)
    console.print("  [green]✓ Token acquired[/green]")

    # Collect governance data per agent
    console.print("\n[bold]Step 2:[/bold] Collecting identity governance data...")
    agents_data = []
    pkgs_data = []

    for agent_name in args.agents:
        console.print(f"  Querying [cyan]{agent_name}[/cyan]...")

        app = find_agent_app(token, agent_name)
        row = {"name": agent_name, "app_id": None}

        if not app:
            console.print(f"    [yellow]⚠ App registration not found — not yet registered in Entra[/yellow]")
            agents_data.append(row)
            continue

        row["app_id"] = app.get("appId")

        # Service principal
        sp = get_service_principal(token, app["appId"])
        row["sp_enabled"] = sp.get("accountEnabled", False) if sp else False
        sp_id = sp.get("id") if sp else None

        # Owners
        owners = get_app_owners(token, app["id"])
        row["owner_count"] = len(owners)

        # Access packages
        pkgs = get_access_packages(token, agent_name)
        row["access_package_count"] = len(pkgs)

        for pkg in pkgs:
            assignments = get_access_package_assignments(token, pkg["id"])
            pkgs_data.append({**pkg, "assignments": assignments})

        # Last sign-in
        signin_activity = None
        if sp_id:
            signin_activity = get_sp_signin_activity(token, sp_id)
        if signin_activity:
            last_signin = signin_activity.get("lastSignInDateTime", "")
            row["last_signin"] = last_signin[:10] if last_signin else "Never"
        else:
            row["last_signin"] = "Never / No data"

        agents_data.append(row)
        time.sleep(0.3)  # Respect Graph throttling

    # Render reports
    console.print("\n[bold]Step 3:[/bold] Identity Governance Report\n")
    render_agent_identity_table(agents_data)

    console.print("\n[bold]Step 4:[/bold] Ownerless Agent Detection\n")
    ownerless = detect_ownerless_agents(token, args.agents)
    render_ownerless_report(ownerless)

    console.print("\n[bold]Step 5:[/bold] Access Package Assignments\n")
    render_access_package_report(pkgs_data)

    # Summary
    registered = sum(1 for d in agents_data if d.get("app_id"))
    governed = sum(
        1 for d in agents_data
        if d.get("owner_count", 0) > 0 and d.get("access_package_count", 0) > 0
    )
    console.print(Panel(
        f"[bold]Governance Summary[/bold]\n"
        f"  Agents queried:    {len(args.agents)}\n"
        f"  Registered in Entra: {registered} / {len(args.agents)}\n"
        f"  Fully governed:    {governed} / {len(args.agents)}\n"
        f"  Ownerless agents:  {len(ownerless)}\n\n"
        f"[dim]Full governance = Owner + Access Package assigned.[/dim]",
        border_style="blue",
        title="[cyan]Results[/cyan]",
    ))


if __name__ == "__main__":
    main()
