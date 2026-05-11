#!/usr/bin/env python3
"""List Foundry project connections to find the AI Search connection id.

Usage:
    set AZURE_AI_PROJECT_ENDPOINT=https://<acct>.services.ai.azure.com/api/projects/<proj>
    az login
    python list_foundry_connections.py
"""

from __future__ import annotations

import os
import sys

from dotenv import load_dotenv
from azure.identity import AzureCliCredential
from azure.ai.projects import AIProjectClient


def main() -> int:
    load_dotenv()
    endpoint = os.environ.get("AZURE_AI_PROJECT_ENDPOINT", "").strip()
    if not endpoint:
        print("ERROR: set AZURE_AI_PROJECT_ENDPOINT", file=sys.stderr)
        return 2

    with (
        AzureCliCredential() as cred,
        AIProjectClient(endpoint=endpoint, credential=cred) as project,
    ):
        print(f"{'NAME':40} {'TYPE':30} ID")
        print("-" * 120)
        for c in project.connections.list():
            print(f"{getattr(c, 'name', ''):40} {str(getattr(c, 'type', '')):30} {getattr(c, 'id', '')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
