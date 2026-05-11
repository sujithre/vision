#!/usr/bin/env python3
"""
setup_foundry_agent_v2.py

One-shot setup: create (or refresh) a Foundry agent wired to the
``chart-Automation`` Azure AI Search index.

This v2 uses azure-ai-agents `AgentsClient` directly, because
azure-ai-projects 2.x removed the classic `project.agents.create_agent(...)`
proxy (its `.agents` is now a versioned-agents API with different methods).

Env (.env auto-loaded):
    AZURE_AI_PROJECT_ENDPOINT   Foundry project endpoint, e.g.
                                https://<resource>.services.ai.azure.com/api/projects/<project>
    MODEL_DEPLOYMENT_NAME       Chat model deployment (e.g. gpt-4o)
    SEARCH_CONNECTION_ID        Full ARM id of the Azure AI Search connection in the
                                Foundry project (Management center -> Connections).
    SEARCH_INDEX                Index name (default: chart-Automation)
    AGENT_NAME                  Agent name (default: chart-automation-agent)
    AGENT_INSTRUCTIONS          Optional: override the system prompt.

Auth: AzureCliCredential -- run ``az login`` first.

Install:
    pip install azure-ai-agents azure-identity python-dotenv
"""

from __future__ import annotations

import os
import sys

from dotenv import load_dotenv
from azure.identity import AzureCliCredential
from azure.ai.agents import AgentsClient
from azure.ai.agents.models import AzureAISearchQueryType, AzureAISearchTool


DEFAULT_INSTRUCTIONS = (
    "You are a presentation analyst. Use the attached Azure AI Search index "
    "(slides indexed from PowerPoint decks with GPT-4o vision summaries) to "
    "answer questions about charts, KPIs, and figures in the decks. Always "
    "cite the deck name and slide number from the search results. If the "
    "index does not contain the answer, say so explicitly rather than "
    "guessing."
)


def _require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        print(f"ERROR: environment variable {name} is required", file=sys.stderr)
        sys.exit(2)
    return value


def main() -> int:
    load_dotenv()

    project_endpoint = _require("AZURE_AI_PROJECT_ENDPOINT")
    model_deployment = _require("MODEL_DEPLOYMENT_NAME")
    search_conn_id = _require("SEARCH_CONNECTION_ID")
    index_name = os.environ.get("SEARCH_INDEX", "chart-Automation").strip()
    agent_name = os.environ.get("AGENT_NAME", "chart-automation-agent").strip()
    instructions = os.environ.get("AGENT_INSTRUCTIONS", "").strip() or DEFAULT_INSTRUCTIONS

    search_tool = AzureAISearchTool(
        index_connection_id=search_conn_id,
        index_name=index_name,
        query_type=AzureAISearchQueryType.VECTOR_SEMANTIC_HYBRID,
        top_k=5,
    )

    with AgentsClient(endpoint=project_endpoint, credential=AzureCliCredential()) as agents:
        # Delete any existing agent with the same name so the new definition
        # (tools, instructions, model) takes effect cleanly.
        for existing in agents.list_agents():
            if getattr(existing, "name", None) == agent_name:
                print(f"Removing existing agent '{agent_name}' (id={existing.id})")
                agents.delete_agent(existing.id)

        agent = agents.create_agent(
            model=model_deployment,
            name=agent_name,
            instructions=instructions,
            tools=search_tool.definitions,
            tool_resources=search_tool.resources,
        )

    print()
    print("Foundry agent ready.")
    print(f"  name : {agent_name}")
    print(f"  id   : {agent.id}")
    print(f"  model: {model_deployment}")
    print(f"  index: {index_name}  (hybrid + semantic, top_k=5)")
    print()
    print("Open Foundry portal -> Agents -> select this agent -> Playground to chat.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
