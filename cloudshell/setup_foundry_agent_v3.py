#!/usr/bin/env python3
"""
setup_foundry_agent_v3.py

Creates (or refreshes) a NEW Foundry portal agent wired to the
``chart-Automation`` Azure AI Search index, using the versioned-agents API
from ``azure-ai-projects`` 2.x. After this runs, the agent shows up in the
new Foundry portal -> Agents tab and is chattable in the Playground.

API used (per official sample
sdk/ai/azure-ai-projects/samples/agents/tools/sample_agent_ai_search.py):
    project_client.agents.create_version(
        agent_name=...,
        definition=PromptAgentDefinition(
            model=..., instructions=...,
            tools=[AzureAISearchTool(azure_ai_search=AzureAISearchToolResource(
                indexes=[AISearchIndexResource(
                    project_connection_id=..., index_name=..., query_type=...
                )]
            ))],
        ),
        description=...,
    )

Env (.env auto-loaded):
    AZURE_AI_PROJECT_ENDPOINT     Foundry project endpoint, e.g.
                                  https://<acct>.services.ai.azure.com/api/projects/<proj>
    MODEL_DEPLOYMENT_NAME         Chat model deployment in the project (e.g. gpt-4o)
    SEARCH_CONNECTION_ID          AI Search **project connection id** as shown in
                                  Foundry portal -> Management center -> Connections
                                  (may be the connection name or full ARM id;
                                  pass exactly what the portal lists).
    SEARCH_INDEX                  Index name (default: chart-Automation)
    AGENT_NAME                    Agent name (default: chart-automation-agent)
    AGENT_DESCRIPTION             Optional description shown in portal.
    AGENT_INSTRUCTIONS            Optional override of the system prompt.

Auth: AzureCliCredential -- run ``az login`` first.

Install:
    pip install "azure-ai-projects>=2.1.0" azure-identity python-dotenv
"""

from __future__ import annotations

import os
import sys

from dotenv import load_dotenv
from azure.identity import AzureCliCredential
from azure.ai.projects import AIProjectClient
from azure.ai.projects.models import (
    AISearchIndexResource,
    AzureAISearchQueryType,
    AzureAISearchTool,
    AzureAISearchToolResource,
    PromptAgentDefinition,
)


DEFAULT_INSTRUCTIONS = (
    "You are a presentation analyst. Use the attached Azure AI Search index "
    "(slides indexed from PowerPoint decks with GPT-4o vision summaries) to "
    "answer questions about charts, KPIs, and figures in the decks. Always "
    "cite the deck name and slide number from the search results. If the "
    "index does not contain the answer, say so explicitly rather than "
    "guessing."
)

DEFAULT_DESCRIPTION = "Grounded Q&A over PowerPoint slides via chart-Automation AI Search index."


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
    description = os.environ.get("AGENT_DESCRIPTION", "").strip() or DEFAULT_DESCRIPTION
    instructions = os.environ.get("AGENT_INSTRUCTIONS", "").strip() or DEFAULT_INSTRUCTIONS

    search_tool = AzureAISearchTool(
        azure_ai_search=AzureAISearchToolResource(
            indexes=[
                AISearchIndexResource(
                    project_connection_id=search_conn_id,
                    index_name=index_name,
                    query_type=AzureAISearchQueryType.SIMPLE,
                )
            ]
        )
    )

    with (
        AzureCliCredential() as credential,
        AIProjectClient(endpoint=project_endpoint, credential=credential) as project_client,
    ):
        # If an agent with this name already exists, delete it so the new
        # definition (tools, instructions, model) takes effect cleanly.
        try:
            existing = project_client.agents.get(agent_name)
            if existing is not None:
                print(f"Removing existing agent '{agent_name}'")
                project_client.agents.delete(agent_name)
        except Exception as ex:  # noqa: BLE001
            # 404 => no existing agent; anything else we surface but continue.
            if "404" not in str(ex) and "ResourceNotFound" not in str(ex):
                print(f"WARN: could not query existing agent '{agent_name}': {ex}",
                      file=sys.stderr)

        agent = project_client.agents.create_version(
            agent_name=agent_name,
            definition=PromptAgentDefinition(
                model=model_deployment,
                instructions=instructions,
                tools=[search_tool],
            ),
            description=description,
        )

    print()
    print("Foundry agent ready.")
    print(f"  name   : {agent.name}")
    print(f"  id     : {agent.id}")
    print(f"  version: {agent.version}")
    print(f"  model  : {model_deployment}")
    print(f"  index  : {index_name}")
    print()
    print("Open the Foundry portal -> Agents -> select this agent -> Playground to chat.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
