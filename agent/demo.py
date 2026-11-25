"""Portfolio Risk Monitor: sidecar demo.

Shows three things side by side, using the Entra ID Auth SDK (sidecar):
  1. /Validate           - the sidecar checks the caller's token
  2. On-behalf-of token  - the agent acts AS THE USER  (delegated, scp claim)
  3. Autonomous token    - the agent acts AS ITSELF    (app-only, roles claim)
The agent never holds a credential. The sidecar does.
"""
import base64
import json
import os
import sys

import msal
import requests
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), "..", ".env"))
TENANT = os.environ["TENANT_ID"]
BLUEPRINT = os.environ["BLUEPRINT_APP_ID"]
AGENT_ID = os.environ["AGENT_IDENTITY_ID"]
FRONTEND = os.environ["FRONTEND_CLIENT_ID"]
SIDECAR = os.getenv("SIDECAR_URL", "http://localhost:5000")


def claims(jwt: str) -> dict:
    """Decode a JWT payload for display. NOT validation."""
    payload = jwt.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    return json.loads(base64.urlsafe_b64decode(payload))


def show(title: str, jwt: str) -> None:
    c = claims(jwt)
    keep = ["aud", "appid", "azp", "oid", "idtyp", "scp", "roles", "upn", "name"]
    print(f"\n--- {title} ---")
    for k in keep:
        if k in c:
            print(f"  {k:6}: {c[k]}")


def user_token() -> str:
    """The 'front end' signs the user in and asks for the agent's access_agent scope."""
    app = msal.PublicClientApplication(FRONTEND, authority=f"https://login.microsoftonline.com/{TENANT}")
    flow = app.initiate_device_flow(scopes=[f"api://{BLUEPRINT}/access_agent"])
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        sys.exit(f"Sign-in failed: {result.get('error_description', result)}")
    return result["access_token"]


def sidecar_header(token: str, **params) -> str:
    r = requests.get(
        f"{SIDECAR}/AuthorizationHeader/Graph",
        params={"AgentIdentity": AGENT_ID, **params},
        headers={"Authorization": f"Bearer {token}"},
        timeout=15,
    )
    if not r.ok:
        sys.exit(f"Sidecar said {r.status_code}: {r.text}")
    return r.json()["authorizationHeader"]  # "Bearer eyJ..."


def graph_get(header: str, path: str) -> None:
    r = requests.get(f"https://graph.microsoft.com/v1.0/{path}", headers={"Authorization": header}, timeout=15)
    print(f"  GET /{path} -> {r.status_code}")


if __name__ == "__main__":
    token = user_token()

    v = requests.get(f"{SIDECAR}/Validate", headers={"Authorization": f"Bearer {token}"}, timeout=15)
    print(f"\n/Validate -> {v.status_code}")

    obo = sidecar_header(token)
    show("1. On behalf of the user (delegated)", obo.split(" ", 1)[1])
    graph_get(obo, "me")

    auto = sidecar_header(token, **{
        "optionsOverride.Scopes": "https://graph.microsoft.com/.default",
        "optionsOverride.RequestAppToken": "true",
    })
    show("2. As the agent itself (autonomous)", auto.split(" ", 1)[1])
    graph_get(auto, "users?$top=1")   # expect 403 until you grant the agent a permission (see README)
