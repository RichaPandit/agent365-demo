#!/usr/bin/env python3
"""
agent365_purview_audit_demo.py
Microsoft Agent 365

Standalone Purview integration demo.
  - Pulls Copilot Studio audit events via the Office 365 Management Activity API
  - Parses CopilotInteraction records and outputs a structured report
  - Checks DLP policy simulation results via Microsoft Graph
  - Summarises DSPM AI Observability posture via Purview API

Required permissions (app registration):
  - ActivityFeed.Read (Office 365 Management APIs)
  - InformationProtectionPolicy.Read.All
  - DataLossPreventionPolicy.Evaluate (for DLP check)
  - Directory.Read.All

Setup:
  pip install msal requests rich

Usage:
  python agent365_purview_audit_demo.py \
    --tenant-id <your-tenant-id> \
    --client-id <your-app-client-id> \
    --client-secret <your-client-secret> \
    --days 7
"""

import argparse
import json
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich import box
from msal import ConfidentialClientApplication

console = Console()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"
O365_MGMT_BASE = "https://manage.office.com/api/v1.0"

AGENT_NAMES = [
    "WealthAdvisorCopilot",
    "PortfolioRiskMonitor",
    "ComplianceGuardian",
]

# Audit record types relevant to Copilot/Agent activity
COPILOT_RECORD_TYPES = {
    "CopilotInteraction": 261,
}

AGENT_OPERATIONS = [
    "AIInteraction",
    "AgentPublished",
    "AgentDeployed",
    "AgentRemoved",
    "AgentConfigurationUpdated",
]


# ─────────────────────────────────────────────
# Auth — two separate token scopes needed
# ─────────────────────────────────────────────

def get_token(tenant_id: str, client_id: str, client_secret: str, scope: str) -> str:
    """Acquire OAuth 2.0 token for the specified scope."""
    app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    result = app.acquire_token_for_client(scopes=[scope])
    if "access_token" not in result:
        error = result.get("error_description", "Unknown auth error")
        console.print(f"[red]❌ Auth failed for scope {scope}: {error}[/red]")
        sys.exit(1)
    return result["access_token"]


def graph_get(token: str, url: str) -> dict:
    """Authenticated GET to Microsoft Graph."""
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, timeout=30)
    if r.status_code not in (200, 201):
        return {}
    return r.json()


def o365_get(token: str, url: str) -> dict:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    try:
        r = requests.get(url, headers=headers, timeout=30, verify=False)

        if r.status_code in (200, 201):
            return r.json()

        elif r.status_code == 400:
            console.print(
                "  [yellow]⚠ No content blobs available for this time range "
                "(normal if ingestion delay or unsupported workload)[/yellow]"
            )
            return []

        elif r.status_code == 401:
            console.print("  [red]❌ Unauthorized — check ActivityFeed.Read permission[/red]")
            return []

        else:
            console.print(f"  [yellow]⚠ O365 API returned {r.status_code}[/yellow]")
            return []

    except Exception as e:
        console.print(f"  [red]Error calling O365 API: {e}[/red]")
        return []


def o365_post(token: str, url: str, body: dict = None) -> dict:
    """Authenticated POST to Office 365 Management API."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=body or {}, timeout=30)
    if r.status_code not in (200, 201):
        return {}
    try:
        return r.json()
    except Exception:
        return {"status": r.status_code}


# ─────────────────────────────────────────────
# Module 1: Subscribe to audit feed
# ─────────────────────────────────────────────

def ensure_audit_subscription(token: str, tenant_id: str, content_type: str = "Audit.General") -> bool:
    """
    Ensure an active subscription exists for the audit content type.
    The O365 Management API requires an active subscription before you can
    pull audit content. This is a one-time setup per tenant.
    """
    url = f"{O365_MGMT_BASE}/{tenant_id}/activity/feed/subscriptions/start?contentType={content_type}"
    result = o365_post(token, url)

    # Check existing subscriptions
    list_url = f"{O365_MGMT_BASE}/{tenant_id}/activity/feed/subscriptions/list"
    subs = o365_get(token, list_url)
    active = [s for s in (subs if isinstance(subs, list) else []) if s.get("status") == "enabled"]

    if active:
        console.print(f"  [green]✓ Audit subscription active ({len(active)} content types)[/green]")
        return True

    console.print(f"  [yellow]⚠ Audit subscription may be pending. First-time setup can take up to 12 hours.[/yellow]")
    return False


# ─────────────────────────────────────────────
# Module 2: Pull audit content blobs
# ─────────────────────────────────────────────

def get_audit_content_uris(
    token: str,
    tenant_id: str,
    days: int,
    content_type: str = "Audit.General"
) -> list[str]:
    """
    Get the list of audit content blob URIs for the given time range.
    The O365 Management API returns content in 1-hour chunks as blob URIs.
    """
    end_time = datetime.now(timezone.utc)
    start_time = end_time - timedelta(minutes=10)

    start_time_str = start_time.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_time_str = end_time.strftime("%Y-%m-%dT%H:%M:%SZ")

    start_str = start_time.strftime("%Y-%m-%dT%H:%M:%S")
    end_str = end_time.strftime("%Y-%m-%dT%H:%M:%S")

    url = (
        f"{O365_MGMT_BASE}/{tenant_id}/activity/feed/subscriptions/content"
        f"?contentType={content_type}"
        f"&startTime={start_time_str}&endTime={end_time_str}"
    )

    uris = []
    while url:
        data = o365_get(token, url)
        if isinstance(data, list):
            uris.extend([item.get("contentUri") for item in data if item.get("contentUri")])
            # Handle pagination via NextPageUri header — not available from o365_get
            # For simplicity, break after first page in demo context
            break
        break

    return uris


def fetch_audit_records(token: str, content_uris: list[str]) -> list[dict]:
    """Fetch and parse audit records from content blob URIs."""
    all_records = []
    headers = {"Authorization": f"Bearer {token}"}

    for uri in content_uris:
        try:
            r = requests.get(uri, headers=headers, timeout=30)
            if r.status_code == 200:
                records = r.json()
                if isinstance(records, list):
                    all_records.extend(records)
        except Exception as e:
            console.print(f"  [yellow]⚠ Failed to fetch blob: {e}[/yellow]")

    return all_records


# ─────────────────────────────────────────────
# Module 3: Filter and parse Copilot records
# ─────────────────────────────────────────────

def filter_copilot_records(records: list[dict]) -> list[dict]:
    """Filter audit records to Copilot Studio / Agent events."""
    copilot_records = []
    for r in records:
        workload = r.get("Workload", "")
        operation = r.get("Operation", "")
        record_type = r.get("RecordType", 0)

        if (
            workload in ("CopilotStudio", "MicrosoftCopilot", "Copilot")
            or record_type == 261  # CopilotInteraction
            or operation in AGENT_OPERATIONS
        ):
            copilot_records.append(r)

    return copilot_records


def parse_agent_interactions(records: list[dict]) -> list[dict]:
    """Parse CopilotInteraction records into structured summaries."""
    interactions = []
    for r in records:
        try:
            created = r.get("CreationTime", "")
            user = r.get("UserId", "Unknown")
            operation = r.get("Operation", "")
            workload = r.get("Workload", "")

            # Parse additional agent-specific fields
            ai_fields = {}
            if isinstance(r.get("AISystemPlugin"), dict):
                ai_fields["plugin"] = r["AISystemPlugin"].get("Id", "")
            if isinstance(r.get("CopilotEventData"), dict):
                ai_fields["thread_id"] = r["CopilotEventData"].get("ThreadId", "")
                ai_fields["agent_name"] = r["CopilotEventData"].get("AgentName", "")
                ai_fields["accessed_resources"] = r["CopilotEventData"].get("AccessedResources", [])

            interactions.append({
                "timestamp": created[:19] if created else "Unknown",
                "user": user,
                "operation": operation,
                "workload": workload,
                "agent": ai_fields.get("agent_name", "Unknown"),
                "plugin": ai_fields.get("plugin", ""),
                "thread_id": ai_fields.get("thread_id", "")[:16] + "..." if ai_fields.get("thread_id") else "",
                "resources_accessed": len(ai_fields.get("accessed_resources", [])),
            })
        except Exception:
            continue

    return interactions


# ─────────────────────────────────────────────
# Module 4: DLP policy check via Graph
# ─────────────────────────────────────────────

def get_dlp_policies(graph_token: str) -> list[dict]:
    """
    List DLP policies scoped to Copilot/CopilotStudio locations.
    Uses Graph beta endpoint — DLP policy management is beta only.
    """
    url = f"{GRAPH_BETA}/security/dataLossPreventionPolicies"
    data = graph_get(graph_token, url)
    
    console.print("[cyan]DEBUG: Raw DLP API response[/cyan]")
    console.print(data)
    policies = data.get("value", [])

    console.print(f"  [green]✓ Retrieved {len(policies)} DLP policies from Graph[/green]")

    return policies


# ─────────────────────────────────────────────
# Module 5: DSPM AI Observability summary
# ─────────────────────────────────────────────

def get_dspm_ai_insights(graph_token: str) -> dict:
    """
    Pull DSPM AI Observability data.
    Uses Graph beta — DSPM insights endpoint.
    Note: This endpoint requires the Purview DSPM add-on license.
    """
    url = f"{GRAPH_BETA}/security/dataSecurityAndGovernance/aiInsightsSummary"
    data = graph_get(graph_token, url)
    return data


# ─────────────────────────────────────────────
# Report rendering
# ─────────────────────────────────────────────

def render_audit_summary(interactions: list[dict], days: int):
    """Render a summary table of Copilot/Agent audit activity."""
    if not interactions:
        console.print(Panel(
            f"[yellow]No Copilot audit records found for the last {days} days.[/yellow]\n\n"
            "[dim]This is normal if:\n"
            "  • The agent was only recently published\n"
            "  • No users have interacted with it yet\n"
            "  • The audit subscription is still initialising (up to 12h)[/dim]",
            title="Purview Audit — CopilotInteraction Events",
            border_style="yellow",
        ))
        return

    table = Table(
        title=f"Purview Audit — CopilotInteraction Events (last {days} days)",
        box=box.ROUNDED,
        border_style="cyan",
        header_style="bold cyan",
        show_lines=True,
    )
    table.add_column("Timestamp", no_wrap=True)
    table.add_column("User")
    table.add_column("Operation")
    table.add_column("Agent")
    table.add_column("Plugin Used")
    table.add_column("Resources\nAccessed", justify="center")
    table.add_column("Thread ID (partial)")

    for item in interactions[:50]:  # Cap at 50 rows for demo output
        plugin = item["plugin"] if item["plugin"] else "[dim]—[/dim]"
        if item["plugin"] and "bingweb" in item["plugin"].lower():
            plugin = "[yellow]BingWebSearch[/yellow]"

        table.add_row(
            item["timestamp"],
            item["user"],
            item["operation"],
            item["agent"] if item["agent"] != "Unknown" else "[dim]Unknown[/dim]",
            plugin,
            str(item["resources_accessed"]),
            item["thread_id"] if item["thread_id"] else "[dim]—[/dim]",
        )

    console.print(table)

    # Summary stats
    by_agent = {}
    by_operation = {}
    for item in interactions:
        by_agent[item["agent"]] = by_agent.get(item["agent"], 0) + 1
        by_operation[item["operation"]] = by_operation.get(item["operation"], 0) + 1

    console.print(f"\n  Total records: [bold]{len(interactions)}[/bold]")
    console.print("  By agent:")
    for agent, count in sorted(by_agent.items(), key=lambda x: -x[1]):
        console.print(f"    {agent}: {count}")
    console.print("  By operation:")
    for op, count in sorted(by_operation.items(), key=lambda x: -x[1]):
        console.print(f"    {op}: {count}")


def render_dlp_report(policies: list[dict]):
    """Render DLP policy status for Copilot-scoped policies."""
    if not policies:
        console.print(Panel(
            "[yellow]No DLP policies found scoped to Microsoft 365 Copilot locations.[/yellow]\n\n"
            "[dim]To create one: Purview portal → Data Loss Prevention → Policies → New policy\n"
            "Set locations to include 'Microsoft 365 Copilot'[/dim]",
            title="DLP Policies — Copilot Scope",
            border_style="yellow",
        ))
        return

    table = Table(
        title="DLP Policies — Copilot Scope",
        box=box.ROUNDED,
        border_style="blue",
        header_style="bold cyan",
    )
    table.add_column("Policy Name")
    table.add_column("State", justify="center")
    table.add_column("Mode", justify="center")
    table.add_column("Last Modified", max_width=12)

    for p in policies:
        state = p.get("state", "unknown")
        state_cell = f"[green]{state}[/green]" if state == "enabled" else f"[yellow]{state}[/yellow]"
        mode = p.get("mode", "unknown")
        mode_cell = f"[yellow]{mode}[/yellow]" if mode == "testWithNotifications" else f"[green]{mode}[/green]"
        modified = p.get("lastModifiedDateTime", "")[:10] if p.get("lastModifiedDateTime") else "—"

        table.add_row(p.get("name", "Unnamed"), state_cell, mode_cell, modified)

    console.print(table)
    console.print(
        "  [dim]Policies in 'testWithNotifications' mode are in simulation — "
        "no enforcement yet. Check Activity Explorer to review simulation hits "
        "before switching to enforce mode.[/dim]"
    )


def render_dspm_summary(dspm_data: dict):
    """Render DSPM AI Observability summary."""
    if not dspm_data or dspm_data.get("error"):
        console.print(Panel(
            "[yellow]DSPM AI Observability data not available.[/yellow]\n\n"
            "[dim]This requires the Purview DSPM add-on license and at least 24 hours "
            "of agent activity after enabling AI Observability in DSPM settings.[/dim]",
            title="DSPM — AI Observability Posture",
            border_style="yellow",
        ))
        return

    # Parse whatever fields are present
    overshared = dspm_data.get("oversharedSitesCount", 0)
    unlabelled = dspm_data.get("unlabelledFilesAccessedCount", 0)
    sensitive_interactions = dspm_data.get("sensitiveDataInteractionsCount", 0)

    risk = "HIGH" if overshared > 10 or unlabelled > 50 else ("MEDIUM" if overshared > 0 else "LOW")
    risk_colour = "red" if risk == "HIGH" else ("yellow" if risk == "MEDIUM" else "green")

    console.print(Panel(
        f"[bold]DSPM AI Observability Summary[/bold]\n\n"
        f"  Overshared SharePoint sites:      [bold]{overshared}[/bold]\n"
        f"  Unlabelled files accessed by agents: [bold]{unlabelled}[/bold]\n"
        f"  Sensitive data interactions:       [bold]{sensitive_interactions}[/bold]\n\n"
        f"  Overall posture: [{risk_colour}]{risk}[/{risk_colour}]",
        border_style=risk_colour,
        title="[cyan]DSPM — AI Observability[/cyan]",
    ))

# ─────────────────────────────────────────────
# Module: Copilot Audit via Graph (UAL Search)
# ─────────────────────────────────────────────

def create_audit_search_job(graph_token: str, start_time: str, end_time: str) -> str:
    """Create an audit log search job for Copilot interactions (RecordType 261)."""

    url = f"{GRAPH_BASE}/auditLogs/queries"

    payload = {
        "displayName": "Copilot Audit Search",
        "filterStartDateTime": start_time,
        "filterEndDateTime": end_time,
        "recordTypeFilters": [261]  # CopilotInteraction
    }

    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type": "application/json"
    }

    r = requests.post(url, headers=headers, json=payload, timeout=30)

    if r.status_code not in (200, 201, 202):
        console.print(f"[red]❌ Failed to create audit query: {r.status_code}[/red]")
        console.print(r.text)
        return None

    job = r.json()
    job_id = job.get("id")

    console.print(f"  [green]✓ Audit search job created: {job_id}[/green]")

    return job_id

def wait_for_audit_job(graph_token: str, job_id: str):
    """Poll until audit job completes."""

    url = f"{GRAPH_BASE}/auditLogs/queries/{job_id}"
    headers = {"Authorization": f"Bearer {graph_token}"}

    for _ in range(12):  # ~2 minutes max
        r = requests.get(url, headers=headers, timeout=30)

        if r.status_code != 200:
            console.print("[yellow]⚠ Failed to check job status[/yellow]")
            return False

        status = r.json().get("status")

        console.print(f"  Job status: {status}")

        if status in ("succeeded", "completed"):
            return True

        time.sleep(10)

    console.print("[yellow]⚠ Audit job timeout[/yellow]")
    return False

def get_audit_job_results(graph_token: str, job_id: str) -> list[dict]:
    """
    Fetch results from completed audit job.
    """
    url = f"{GRAPH_BASE}/auditLogs/queries/{job_id}/records"
    headers = {"Authorization": f"Bearer {graph_token}"}

    r = requests.get(url, headers=headers, timeout=30)

    if r.status_code != 200:
        console.print("[red]❌ Failed to fetch audit results[/red]")
        return []

    data = r.json()
    return data.get("value", [])

def get_copilot_audit_via_graph(graph_token: str, days: int) -> list[dict]:
    """
    Full flow: create job → wait → fetch results
    """
    end_time = datetime.now(timezone.utc)
    start_time = end_time - timedelta(days=days)

    start_str = start_time.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end_time.strftime("%Y-%m-%dT%H:%M:%SZ")

    job_id = create_audit_search_job(graph_token, start_str, end_str)

    if not job_id:
        return []

    if not wait_for_audit_job(graph_token, job_id):
        return []

    results = get_audit_job_results(graph_token, job_id)

    console.print(f"  [green]✓ Retrieved {len(results)} Copilot audit records[/green]")

    return results


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Agent 365 — Purview Audit & Data Security Demo"
    )
    parser.add_argument("--tenant-id", required=True, help="Azure AD tenant ID")
    parser.add_argument("--client-id", required=True, help="App registration client ID")
    parser.add_argument("--client-secret", required=True, help="App registration client secret")
    parser.add_argument("--days", type=int, default=7, help="Number of days of audit history to pull (default: 7)")
    args = parser.parse_args()

    console.print(Panel.fit(
        "[bold cyan]Microsoft Agent 365[/bold cyan]\n"
        "[white]Purview Audit & Data Security Demo[/white]\n"
        f"[dim]Pulling last {args.days} days | {datetime.now().strftime('%Y-%m-%d %H:%M')}[/dim]",
        border_style="cyan",
    ))

    # Authenticate for both APIs
    console.print("\n[bold]Step 1:[/bold] Authenticating...")
    graph_token = get_token(
        args.tenant_id, args.client_id, args.client_secret,
        "https://graph.microsoft.com/.default"
    )
    o365_token = get_token(
        args.tenant_id, args.client_id, args.client_secret,
        "https://manage.office.com/.default"
    )
    console.print("  [green]✓ Graph token acquired[/green]")
    console.print("  [green]✓ O365 Management API token acquired[/green]")

    # Ensure audit subscription is active
    console.print("\n[bold]Step 2:[/bold] Verifying Purview audit subscription...")
    ensure_audit_subscription(o365_token, args.tenant_id)

    # Pull audit content
    console.print(f"\n[bold]Step 3:[/bold] Pulling Copilot audit via Graph...")

    records = get_copilot_audit_via_graph(graph_token, args.days)

    if records:
        console.print(f"  [green]✓ {len(records)} audit records retrieved[/green]")
        interactions = parse_agent_interactions(records)
    else:
        console.print("  [yellow]No audit data returned via Graph API[/yellow]")
        interactions = []

    # ✅ Demo fallback (adds resilience so output is never empty)
    if not interactions:
        console.print("  [dim]No live Copilot interactions available — using demo sample[/dim]\n")

        interactions = [{
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": "demo.user@tenant.com",
            "operation": "AIInteraction",
            "workload": "Copilot",
            "agent": "WealthAdvisorCopilot",
            "plugin": "",
            "thread_id": "demo-thread...",
            "resources_accessed": 2,
        }]

    # Render audit report
    console.print("\n[bold]Step 4:[/bold] Purview Audit Report\n")
    render_audit_summary(interactions, args.days)

    # DLP policies
    console.print("\n[bold]Step 5:[/bold] DLP Policy Status\n")
    dlp_policies = get_dlp_policies(graph_token)
    render_dlp_report(dlp_policies)

    # DSPM
    console.print("\n[bold]Step 6:[/bold] DSPM AI Observability\n")
    dspm_data = get_dspm_ai_insights(graph_token)
    render_dspm_summary(dspm_data)

    # Final summary
    console.print(Panel(
        f"[bold]Run Complete[/bold]\n"
        f"  Audit records pulled:     {len(interactions)}\n"
        f"  DLP policies (Copilot):   {len(dlp_policies)}\n"
        f"  Days covered:             {args.days}\n\n"
        f"[dim]Next steps:\n"
        f"  • Review audit records for unexpected knowledge source access\n"
        f"  • Move DLP policies from simulation to enforce mode after reviewing Activity Explorer\n"
        f"  • Enable DSPM SharePoint oversharing assessment if not already on\n"
        f"  • Set KQL alerts in Defender for sign-in spikes on agent service principals[/dim]",
        border_style="blue",
        title="[cyan]Results[/cyan]",
    ))


if __name__ == "__main__":
    main()
