# Microsoft Agent 365 — Investment & Wealth Management Demo
> **Giving AI Agents Their Identity in a Regulated Finance Environment**
> 
> A Day 1 hands-on demo showing how to use Microsoft Agent 365 and the
> Microsoft Graph API to discover and register AI agent identities in an
> Investment & Wealth Management context — with Zero Trust principles
> applied from the very first line of code.

-----

## The Three Agents in This Demo

|Agent                     |Purpose                                              |Risk Level|Compliance            |

|--------------------------|-----------------------------------------------------|----------|----------------------|

|🛡️ **PortfolioRiskMonitor**|Real-time portfolio risk monitoring & alerting       |🔴 HIGH    |MiFID II, FCA COBS    |

|🤝 **WealthAdvisorCopilot**|Investment advice assistant for relationship managers|🟡 MEDIUM  |FCA COBS, GDPR        |

|⚖️ **ComplianceGuardian**  |Regulatory compliance monitoring & audit reporting   |🔴 HIGH    |MiFID II, AML, FCA SUP|

These are realistic agent archetypes you would find in any wealth management firm deploying AI today — each with different risk profiles, data access patterns, and compliance obligations.

-----

## What This Demo Does

|Command                                  |What it shows                                                                                                       |

|-----------------------------------------|--------------------------------------------------------------------------------------------------------------------|

|`discover`                               |Lists all AI agent identities registered in your Entra tenant with a risk snapshot — *“who is already on the team?”*|

|`register --agent portfolio-risk-monitor`|Registers PortfolioRiskMonitor with HIGH risk flags and compliance-aware next steps                                 |

|`register --agent wealth-advisor`        |Registers WealthAdvisorCopilot with MEDIUM risk profile                                                             |

|`register --agent compliance-guardian`   |Registers ComplianceGuardian with HIGH risk flags and RegTech tagging                                               |

|`register --agent custom --name "X"`     |Register any agent of your own                                                                                      |

-----

## Prerequisites

- Python 3.9+

- An Azure subscription with Entra ID

- An App Registration with these **API permissions** (Application type, admin consented):

  - `Application.Read.All` — for discovery

  - `Application.ReadWrite.All` — for registration

-----

## Setup

### 1. Clone and install

```bash

git clone https://github.com/YOUR_USERNAME/agent365-finance-demo.git

cd agent365-finance-demo

python -m venv venv

source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt

```

### 2. Configure credentials

```bash

cp .env.example .env

# Edit .env with your Azure tenant ID, client ID, and client secret

```

-----

## Usage

### Discover all agents in your tenant

```bash

python agent365_finance_demo.py discover

```

Filter by keyword:

```bash

python agent365_finance_demo.py discover --filter sentinel

python agent365_finance_demo.py discover --filter compliance

```

**Sample output:**

```

━━━  MODULE 1: Agent Discovery & Inventory  ━━━

╭─────────────────────────────────────┬─────────────┬──────────┬────────────┬──────────┬──────────────────────────────╮

│ Agent Name                          │ Type        │ App ID   │ Registered │ Status   │ Tags                         │

├─────────────────────────────────────┼─────────────┼──────────┼────────────┼──────────┼──────────────────────────────┤

│ PortfolioRiskMonitor-WealthMgmt        │ Application │ a1b2c3.. │ 2025-04-01 │ ● Active │ WealthManagement, AIAgent... │

│ WealthAdvisorCopilot-Client         │ Application │ d4e5f6.. │ 2025-04-15 │ ● Active │ WealthManagement, AIAgent... │

│ ComplianceGuardian-RegTech          │ Application │ g7h8i9.. │ 2025-05-01 │ ● Active │ Compliance, RegTech, AIAgent │

╰─────────────────────────────────────┴─────────────┴──────────┴────────────┴──────────┴──────────────────────────────╯

  📊 Wealth Management Agent Inventory — Risk Snapshot

  ┌──────────────────────────────────────────────────

  │  Total agents discovered : 3

  │  Active                  : 3

  │  Disabled / Orphaned     : 0

  │  Governance tagged       : 3

  │  Untagged (blind spots)  : 0

  └──────────────────────────────────────────────────

```

-----

### Register a finance agent

```bash

# Register the portfolio risk monitoring agent

python agent365_finance_demo.py register --agent portfolio-risk-monitor

# Register the client advisory assistant

python agent365_finance_demo.py register --agent wealth-advisor

# Register the compliance & audit agent

python agent365_finance_demo.py register --agent compliance-guardian

```

**Sample output for PortfolioRiskMonitor (HIGH risk):**

```

  ┌─ Agent Profile ────────────────────────────────────────────────

  │  📊  PortfolioRiskMonitor-WealthMgmt

  │  Purpose    : Real-time portfolio risk monitoring & alerting

  │  Risk Level : HIGH

  │  Data Access: Portfolio DB, Market Data Feed, Client Risk Profiles

  │  Compliance : MiFID II, FCA COBS

  │  Tags       : WealthManagement, RiskMonitoring, AIAgent, Agent365Demo

  └────────────────────────────────────────────────────────────────

[✓] App Registration created.

[✓] Service Principal created.

  🔐 Zero Trust Identity Summary

  ┌────────────────────────────────────────────────────────────

  │  📊  Agent       : PortfolioRiskMonitor-WealthMgmt-20250501...

  │  🔒  Scope       : AzureADMyOrg (tenant-only, least privilege)

  │  ⚠️   Risk Level  : HIGH

  │  📋  Compliance  : MiFID II, FCA COBS

  └────────────────────────────────────────────────────────────

  📋 Recommended Next Steps:

     ⚠️  HIGH RISK AGENT — additional controls required:

     1. Assign only the minimum Graph API permissions needed

     2. Enable Conditional Access policy for this agent identity

     3. Configure Microsoft Defender alerts for behavioral anomalies

     ...

  Your AI companion is ready. Introduce it to Agent 365. 🤝

```

-----

## Zero Trust Principles Applied

Every agent registered by this demo applies Zero Trust from line one:

- **Verify Explicitly** → unique Entra identity per agent, no shared credentials

- **Least Privilege** → `AzureADMyOrg` audience, no standing permissions granted

- **Assume Breach** → risk level flags, compliance tagging, JSON audit trail saved locally

-----

## Project Structure

```

agent365-finance-demo/

├── agent365_finance_demo.py   # Main CLI script

├── requirements.txt           # Python dependencies

├── .env.example               # Credential template (safe to commit)

├── .gitignore                 # Keeps .env and JSON output out of git

└── README.md                  # This file

```

-----

## Related Blog Post

**[Microsoft Agent 365: Giving Your AI Companion the Identity It Deserves](#)**

-----

## License

MIT