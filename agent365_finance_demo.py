#!/usr/bin/env python3
"""
Microsoft Agent 365 — Investment & Wealth Management Demo
----------------------------------------------------------

  DISCOVER  — Inventory the 3 AI agents already running in your
              wealth management environment
  REGISTER  — Onboard a brand new agent identity with Zero Trust
              principles applied from day one

Agents in this demo:
  1. PortfolioRiskMonitor    — Real-time portfolio risk monitoring agent
  2. WealthAdvisorCopilot — Client-facing investment advice assistant
  3. ComplianceGuardian   — Regulatory compliance & audit agent

Usage:
  python agent365_finance_demo.py discover
  python agent365_finance_demo.py discover --filter sentinel
  python agent365_finance_demo.py register --agent portfolio-risk-monitor
  python agent365_finance_demo.py register --agent wealth-advisor
  python agent365_finance_demo.py register --agent compliance-guardian
  python agent365_finance_demo.py register --agent custom --name "MyAgent" --description "My custom agent"
"""

import os
import sys
import json
import argparse
from datetime import datetime, timezone

import time
import requests
from dotenv import load_dotenv
from tabulate import tabulate
from colorama import init, Fore, Style

# ── Init ──────────────────────────────────────────────────────────────────────
init(autoreset=True)
load_dotenv()

TENANT_ID     = os.getenv("AZURE_TENANT_ID")
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL  = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

# ── Finance Agent Profiles ────────────────────────────────────────────────────
#
# These are the three agents in our Investment & Wealth Management scenario.
# Each has a realistic name, purpose, risk profile, and tags that would appear
# in a real Agent 365 inventory.
#

FINANCE_AGENTS = {
    "portfolio-risk-monitor": {
        "displayName":   "PortfolioRiskMonitor-WealthMgmt",
        "description":   (
            "Real-time portfolio risk monitoring agent. Continuously evaluates "
            "client portfolios against market signals, VaR thresholds, and "
            "concentration limits. Triggers alerts when risk exposure exceeds "
            "agreed parameters."
        ),
        "tags":          ["WealthManagement", "RiskMonitoring", "AIAgent", "Agent365Demo"],
        "purpose":       "Real-time portfolio risk monitoring & alerting",
        "risk_level":    "HIGH",          # high — touches live portfolio data
        "data_access":   ["Portfolio DB", "Market Data Feed", "Client Risk Profiles"],
        "compliance":    "MiFID II, FCA COBS",
        "emoji":         "📊",
    },
    "wealth-advisor": {
        "displayName":   "WealthAdvisorCopilot-Client",
        "description":   (
            "Client-facing AI assistant that helps relationship managers prepare "
            "personalised investment recommendations, draft client communications, "
            "and summarise portfolio performance reports. Operates under human "
            "supervision before any client-facing output is sent."
        ),
        "tags":          ["WealthManagement", "ClientAdvisory", "AIAgent", "Agent365Demo"],
        "purpose":       "Investment advice assistant for relationship managers",
        "risk_level":    "MEDIUM",
        "data_access":   ["CRM", "Portfolio Reports", "Research Library"],
        "compliance":    "FCA COBS, GDPR",
        "emoji":         "🤝",
    },
    "compliance-guardian": {
        "displayName":   "ComplianceGuardian-RegTech",
        "description":   (
            "Automated regulatory compliance and audit agent. Monitors trade "
            "activity for suspicious patterns, checks transactions against "
            "sanctions lists, generates regulatory reports, and flags potential "
            "breaches for human review. Runs on a scheduled cadence and on-demand."
        ),
        "tags":          ["WealthManagement", "Compliance", "RegTech", "AIAgent", "Agent365Demo"],
        "purpose":       "Regulatory compliance monitoring & audit reporting",
        "risk_level":    "HIGH",          # high — reads all transaction data
        "data_access":   ["Transaction Ledger", "Sanctions DB", "Audit Logs", "Trade Blotter"],
        "compliance":    "MiFID II, AML, FCA SUP",
        "emoji":         "⚖️",
    },
}

RISK_COLOR = {
    "HIGH":   Fore.RED,
    "MEDIUM": Fore.YELLOW,
    "LOW":    Fore.GREEN,
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def banner():
    print(Fore.CYAN + Style.BRIGHT + """
╔══════════════════════════════════════════════════════════════════╗
║     Microsoft Agent 365 — Investment & Wealth Management Demo    ║
║          Agent Discovery & Identity Registration                 ║
╚══════════════════════════════════════════════════════════════════╝
""")


def check_env():
    missing = [v for v in ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"]
               if not os.getenv(v)]
    if missing:
        print(Fore.RED + f"\n[ERROR] Missing environment variables: {', '.join(missing)}")
        print(Fore.YELLOW + "  → Copy .env.example to .env and fill in your Azure credentials.\n")
        sys.exit(1)


def get_access_token() -> str:
    print(Fore.BLUE + "[*] Authenticating with Microsoft Entra ID...")
    resp = requests.post(TOKEN_URL, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=15)
    if resp.status_code != 200:
        print(Fore.RED + f"[ERROR] Auth failed ({resp.status_code}): {resp.text}")
        sys.exit(1)
    print(Fore.GREEN + "[✓] Authenticated successfully.\n")
    return resp.json()["access_token"]

def graph_get(token: str, path: str, params: dict = None) -> dict:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.get(f"{GRAPH_BASE}{path}", headers=headers, params=params, timeout=15)
    if resp.status_code not in (200, 201):
        print(Fore.RED + f"[ERROR] Graph API ({resp.status_code}): {resp.text}")
        sys.exit(1)
    return resp.json()

def graph_post(token: str, path: str, payload: dict) -> dict:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(f"{GRAPH_BASE}{path}", headers=headers, json=payload, timeout=15)
    if resp.status_code not in (200, 201):
        print(Fore.RED + f"[ERROR] Graph API ({resp.status_code}): {resp.text}")
        sys.exit(1)
    return resp.json()

def print_agent_card(profile: dict):
    """Pretty-print a finance agent profile before registration."""
    rc = RISK_COLOR.get(profile["risk_level"], Fore.WHITE)
    print(Fore.CYAN + "  ┌─ Agent Profile " + "─" * 48)
    print(f"  │  {profile['emoji']}  {Fore.WHITE + Style.BRIGHT}{profile['displayName']}")
    print(f"  │  Purpose    : {profile['purpose']}")
    print(f"  │  Risk Level : {rc}{profile['risk_level']}{Style.RESET_ALL}")
    print(f"  │  Data Access: {', '.join(profile['data_access'])}")
    print(f"  │  Compliance : {profile['compliance']}")
    print(f"  │  Tags       : {', '.join(profile['tags'])}")
    print(Fore.CYAN + "  └" + "─" * 64 + "\n")

# ── Feature 1: Discover & Inventory ──────────────────────────────────────────

def discover_agents(token: str, keyword: str = None):
    print(Fore.CYAN + Style.BRIGHT + "━━━  MODULE 1: Agent Discovery & Inventory  ━━━\n")
    print(Fore.WHITE + "  Scanning your Entra tenant for registered AI agent identities...\n")

    params = {
        "$select": "id,displayName,appId,servicePrincipalType,tags,createdDateTime,accountEnabled",
        "$top": 999,
    }
    data   = graph_get(token, "/servicePrincipals", params=params)
    agents = data.get("value", [])

    # ── Filter to Agent365Demo tagged principals only (or by keyword) ─────────
    demo_agents = [
        a for a in agents
        if "Agent365Demo" in (a.get("tags") or [])
        or "AIAgent" in (a.get("tags") or [])
    ]

    if keyword:
        kw = keyword.lower()
        demo_agents = [
            a for a in demo_agents
            if kw in (a.get("displayName") or "").lower()
            or any(kw in t.lower() for t in (a.get("tags") or []))
        ]

    if not demo_agents:
        print(Fore.YELLOW + "  No AI agents found" +
              (f" matching '{keyword}'." if keyword else "."))
        print(Fore.YELLOW + "  → Run the register command to onboard your first agent.\n")
        return

    # ── Build display table ───────────────────────────────────────────────────
    rows = []
    for a in demo_agents:
        name    = a.get("displayName", "—")
        created = (a.get("createdDateTime") or "")[:10] or "—"
        status  = (Fore.GREEN + "● Active" if a.get("accountEnabled")
                   else Fore.RED + "● Disabled") + Style.RESET_ALL
        tags    = ", ".join(a.get("tags") or []) or "—"
        sp_type = a.get("servicePrincipalType", "—")
        app_id  = (a.get("appId") or "—")[:8] + "..."
        rows.append([name, sp_type, app_id, created, status, tags[:45] + ("…" if len(tags) > 45 else "")])

    headers = ["Agent Name", "Type", "App ID", "Registered", "Status", "Tags"]
    print(tabulate(rows, headers=headers, tablefmt="rounded_outline"))

    # ── Risk snapshot ─────────────────────────────────────────────────────────
    active   = sum(1 for a in demo_agents if a.get("accountEnabled"))
    disabled = len(demo_agents) - active
    tagged   = sum(1 for a in demo_agents if a.get("tags"))

    print()
    print(Fore.CYAN + Style.BRIGHT + "  📊 Wealth Management Agent Inventory — Risk Snapshot")
    print(Fore.WHITE + f"  ┌{'─'*50}")
    print(f"  │  Total agents discovered : {len(demo_agents)}")
    print(f"  │  Active                  : {Fore.GREEN}{active}{Style.RESET_ALL}")
    print(f"  │  Disabled / Orphaned     : {Fore.RED}{disabled}{Style.RESET_ALL}")
    print(f"  │  Governance tagged       : {Fore.YELLOW}{tagged}{Style.RESET_ALL}")
    print(f"  │  Untagged (blind spots)  : {Fore.RED}{len(demo_agents) - tagged}{Style.RESET_ALL}")
    print(Fore.WHITE + f"  └{'─'*50}\n")

    # ── Zero Trust advisory ───────────────────────────────────────────────────
    if disabled > 0:
        print(Fore.YELLOW + "  ⚠️  Zero Trust Advisory:")
        print("     Disabled agents should be decommissioned from Entra.")
        print("     Orphaned identities are a standing attack surface.\n")

    # ── Save inventory ────────────────────────────────────────────────────────
    fname = "wealthmgmt_agent_inventory.json"
    with open(fname, "w") as f:
        json.dump(demo_agents, f, indent=2)
    print(Fore.GREEN + f"  [✓] Full inventory saved to {fname}\n")


# ── Feature 2: Register Agent Identity ───────────────────────────────────────

def register_agent(token: str, agent_key: str,
                   custom_name: str = None, custom_desc: str = None):
    print(Fore.CYAN + Style.BRIGHT + "━━━  MODULE 2: Agent Identity Registration  ━━━\n")

    # ── Resolve profile ───────────────────────────────────────────────────────
    if agent_key == "custom":
        if not custom_name:
            print(Fore.RED + "[ERROR] --name is required for custom agents.")
            sys.exit(1)
        profile = {
            "displayName": custom_name,
            "description": custom_desc or f"Custom AI agent registered via Agent365 Demo.",
            "tags":        ["AIAgent", "Agent365Demo", "WealthManagement"],
            "purpose":     custom_desc or "Custom agent",
            "risk_level":  "MEDIUM",
            "data_access": ["TBD — assign permissions explicitly"],
            "compliance":  "Review required",
            "emoji":       "🤖",
        }
    else:
        profile = FINANCE_AGENTS.get(agent_key)
        if not profile:
            print(Fore.RED + f"[ERROR] Unknown agent key: '{agent_key}'")
            print(Fore.YELLOW + f"  Valid options: {', '.join(FINANCE_AGENTS.keys())}, custom")
            sys.exit(1)

    # ── Show profile ──────────────────────────────────────────────────────────
    print(Fore.WHITE + "  Registering the following agent identity:\n")
    print_agent_card(profile)

    # ── Step 1: App Registration ──────────────────────────────────────────────
    ts       = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    app_name = f"{profile['displayName']}-{ts}"

    print(Fore.BLUE + f"[*] Creating App Registration: {app_name}")
    app = graph_post(token, "/applications", {
        "displayName":    app_name,
        "description":    profile["description"],
        "tags":           profile["tags"],
        "signInAudience": "AzureADMyOrg",
        "notes":          (
            f"Agent365 Demo | Domain: Investment & Wealth Management | "
            f"Risk: {profile['risk_level']} | Registered: {ts}"
        ),
    })
    app_id     = app["appId"]
    app_obj_id = app["id"]
    print(Fore.GREEN + "[✓] App Registration created.")
    print(f"     Display Name : {app['displayName']}")
    print(f"     App ID       : {app_id}")
    print(f"     Object ID    : {app_obj_id}\n")

    # ── Step 2: Service Principal ─────────────────────────────────────────────
    print(Fore.BLUE + "[*] Creating Service Principal (agent identity)...")
    print(Fore.YELLOW + "  [~] Waiting for Entra to propogate app registration...")
    sp_payload = {
        "appId":          app_id,
        "tags":           profile["tags"] + ["WindowsAzureActiveDirectoryIntegratedApp"],
        "accountEnabled": True,
    }
    sp = None
    for attempt in range(1, 6):
        time.sleep(5)
        try:
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            resp = requests.post(f"{GRAPH_BASE}/servicePrincipals", headers=headers, json=sp_payload, timeout=15)
            if resp.status_code in (200, 201):
                sp = resp.json()
                break
            elif resp.status_code == 400 and "NoBackingApplicationObject" in resp.text:
                print(Fore.YELLOW + f"  [~] Attempt {attempt}/5 - still propagating, retrying...")
            else:
                print(Fore.RED + f"[ERROR] Graph API ({resp.status_code}): {resp.text}")
                sys.exit(1)
        except Exception as e:
            print(Fore.RED + f"[ERROR] Request failed: {e}")
            sys.exit(1)

    if not sp:
        print(Fore.RED + "[ERROR] Service Principal creation failed after 5 attempts.")
        print(Fore.YELLOW + "  → App Registration was created successfully.")
        print(Fore.YELLOW + "  → Wait 30 seconds and re-run register command.")
        sys.exit(1)
        
    sp_id = sp["id"]
    print(Fore.GREEN + "[✓] Service Principal created.")
    print(f"     SP Object ID : {sp_id}")
    print(f"     Type         : {sp.get('servicePrincipalType', '—')}")
    print(f"     Enabled      : {sp.get('accountEnabled')}\n")

    # ── Step 3: Zero Trust Summary ────────────────────────────────────────────
    rc = RISK_COLOR.get(profile["risk_level"], Fore.WHITE)
    print(Fore.CYAN + Style.BRIGHT + "  🔐 Zero Trust Identity Summary")
    print(Fore.WHITE + f"  ┌{'─'*60}")
    print(f"  │  {profile['emoji']}  Agent          : {app_name}")
    print(f"  │  🏢  Tenant          : {TENANT_ID}")
    print(f"  │  🪪  App ID          : {app_id}")
    print(f"  │  🔑  SP Object ID    : {sp_id}")
    print(f"  │  🔒  Sign-in Scope   : AzureADMyOrg (tenant-only, least privilege)")
    print(f"  │  ⚠️   Risk Level      : {rc}{profile['risk_level']}{Style.RESET_ALL}")
    print(f"  │  📋  Compliance      : {profile['compliance']}")
    print(f"  │  🕒  Registered at   : {datetime.now(timezone.utc).isoformat()}")
    print(Fore.WHITE + f"  └{'─'*60}\n")

    # ── Step 4: Next steps based on risk level ────────────────────────────────
    print(Fore.YELLOW + "  📋 Recommended Next Steps:")
    if profile["risk_level"] == "HIGH":
        print(Fore.RED + "     ⚠️  HIGH RISK AGENT — additional controls required:")
        print("     1. Assign only the minimum Graph API permissions needed")
        print("     2. Enable Conditional Access policy for this agent identity")
        print("     3. Configure Microsoft Defender alerts for behavioral anomalies")
        print("     4. Register in Agent 365 for real-time observability monitoring")
        print("     5. Schedule quarterly access reviews in Entra ID Governance")
        print("     6. Notify your compliance team — this agent touches regulated data\n")
    else:
        print("     1. Assign only the minimum Graph API permissions needed")
        print("     2. Enable Conditional Access policy for this agent identity")
        print("     3. Register in Agent 365 for observability monitoring")
        print("     4. Schedule bi-annual access reviews in Entra ID Governance\n")

    # ── Step 5: Save registration record ─────────────────────────────────────
    record = {
        "agent_name":     app_name,
        "domain":         "Investment & Wealth Management",
        "purpose":        profile["purpose"],
        "risk_level":     profile["risk_level"],
        "app_id":         app_id,
        "app_object_id":  app_obj_id,
        "sp_object_id":   sp_id,
        "tenant_id":      TENANT_ID,
        "registered_at":  datetime.now(timezone.utc).isoformat(),
        "tags":           profile["tags"],
        "data_access":    profile["data_access"],
        "compliance":     profile["compliance"],
        "zero_trust": {
            "sign_in_audience": "AzureADMyOrg",
            "principle":        "Verify Explicitly — Least Privilege — Assume Breach",
        },
    }
    fname = f"registered_{agent_key}_{ts}.json"
    with open(fname, "w") as f:
        json.dump(record, f, indent=2)
    print(Fore.GREEN + f"  [✓] Registration record saved to {fname}")
    print(Fore.CYAN  + f"\n  Your AI companion is ready. Introduce it to Agent 365. 🤝\n")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    banner()
    check_env()

    parser = argparse.ArgumentParser(
        description="Agent 365 Day 1 Demo — Investment & Wealth Management"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # discover
    d = sub.add_parser("discover", help="Inventory AI agents in your Entra tenant")
    d.add_argument("--filter", metavar="KEYWORD",
                   help="Filter by name/tag keyword (e.g. sentinel, compliance)")

    # register
    r = sub.add_parser("register", help="Register a new agent identity in Entra ID")
    r.add_argument("--agent", required=True,
                   choices=list(FINANCE_AGENTS.keys()) + ["custom"],
                   help="Agent to register (or 'custom' for your own)")
    r.add_argument("--name",        help="Name for custom agent")
    r.add_argument("--description", help="Description for custom agent")

    args  = parser.parse_args()
    token = get_access_token()

    if args.command == "discover":
        discover_agents(token, keyword=args.filter)
    elif args.command == "register":
        register_agent(token,
                       agent_key=args.agent,
                       custom_name=getattr(args, "name", None),
                       custom_desc=getattr(args, "description", None))

if __name__ == "__main__":
    main()