#!/usr/bin/env python3
"""
Create a Microsoft Foundry agent (via the Microsoft Agent Framework) and wire
the Azure AI Search index produced by ``ingest_vision_to_search.py`` in as a
knowledge tool, then chat with it.

Pipeline:
    1) analyze_pptx_vision.py   -> deck.vision.slides.jsonl
    2) ingest_vision_to_search.py -> chart-Automation index in Azure AI Search
    3) agent_with_search.py     <-- this script: agent + AI Search tool + chat

Env (.env auto-loaded):
    AZURE_AI_PROJECT_ENDPOINT   Foundry project endpoint
                                (e.g. https://<resource>.services.ai.azure.com/api/projects/<project>)
    MODEL_DEPLOYMENT_NAME       Chat model deployment in the same Foundry project
                                (e.g. gpt-4o, gpt-4.1)
    SEARCH_CONNECTION_ID        Full ARM id of the Azure AI Search connection in
                                the Foundry project. Easiest way:
                                  az ml connection list -g <rg> --workspace-name <project> -o table
                                or copy from Foundry portal -> Management center -> Connections.
                                Format:
                                  /subscriptions/<sub>/resourceGroups/<rg>/providers/
                                  Microsoft.CognitiveServices/accounts/<account>/projects/<project>/
                                  connections/<connection-name>
    SEARCH_INDEX                Index name (default: chart-Automation)
    AGENT_NAME                  Agent name (default: chart-automation-agent)
    AGENT_KEEP                  If "1" leave the agent in the project after exit
                                (default: deletes it so re-runs stay clean)

Auth: DefaultAzureCredential / AzureCliCredential (run `az login` first).

Required RBAC on the Foundry project:
  - Azure AI User (or Azure AI Project Manager)
The Search connection itself must already exist in the Foundry project and
point at the same Search service that hosts ``chart-Automation``.

Install:
    pip install agent-framework azure-ai-agents azure-identity python-dotenv
"""

from __future__ import annotations

import asyncio
import os
import sys
from typing import Optional

from dotenv import load_dotenv

from azure.identity.aio import AzureCliCredential
from azure.ai.agents.models import AzureAISearchTool, AzureAISearchQueryType

from agent_framework.azure import AzureAIAgentClient


load_dotenv()

PROJECT_ENDPOINT = os.environ.get("AZURE_AI_PROJECT_ENDPOINT", "").rstrip("/")
MODEL_DEPLOYMENT = os.environ.get("MODEL_DEPLOYMENT_NAME", "")
SEARCH_CONN_ID = os.environ.get("SEARCH_CONNECTION_ID", "")
SEARCH_INDEX = os.environ.get("SEARCH_INDEX", "chart-Automation")
AGENT_NAME = os.environ.get("AGENT_NAME", "chart-automation-agent")
KEEP_AGENT = os.environ.get("AGENT_KEEP", "0") == "1"

INSTRUCTIONS = """\
You are the Chart-Automation analyst. You answer questions about a slide deck
that has been indexed in Azure AI Search (index: chart-Automation). Each search
hit represents one slide and includes:
  - slide_title, slide_summary, slide_description
  - highlighted_terms (bold/colored text from the slide)
  - kpis_text, charts_text, tables_text, callouts_text
  - content (the full slide markdown produced by GPT-4o vision)

Rules:
- ALWAYS retrieve from the AI Search tool before answering content questions.
- Cite slides as "slide N - <slide_title>" using the slide_number field.
- Prefer numbers from kpis_text and charts_text when quantifying anything.
- If no slide is relevant, say so explicitly. Do NOT invent figures.
- Be concise: bullets, ~5 lines unless the user asks for detail.
"""

DEMO_QUESTIONS = [
    "Summarize the deck in 5 bullets and list the slides each bullet draws from.",
    "Which slides talk about chart automation priorities? Quote the KPIs.",
    "What are the highlighted terms across the deck and what do they group into?",
]


def _require(name: str, value: str) -> None:
    if not value:
        print(f"ERROR: env var {name} is required (set it in .env)", file=sys.stderr)
        sys.exit(2)


async def run() -> int:
    _require("AZURE_AI_PROJECT_ENDPOINT", PROJECT_ENDPOINT)
    _require("MODEL_DEPLOYMENT_NAME", MODEL_DEPLOYMENT)
    _require("SEARCH_CONNECTION_ID", SEARCH_CONN_ID)

    print(f"Project    : {PROJECT_ENDPOINT}")
    print(f"Model      : {MODEL_DEPLOYMENT}")
    print(f"Index      : {SEARCH_INDEX}")
    print(f"Connection : {SEARCH_CONN_ID.rsplit('/', 1)[-1]}")
    print(f"Agent      : {AGENT_NAME} (keep={KEEP_AGENT})\n")

    # AzureAISearchTool surfaces our hybrid index to the agent.
    # VECTOR_SEMANTIC_HYBRID = BM25 + vector + L2 semantic ranker, which matches
    # what we built in ingest_vision_to_search.py.
    ai_search = AzureAISearchTool(
        index_connection_id=SEARCH_CONN_ID,
        index_name=SEARCH_INDEX,
        query_type=AzureAISearchQueryType.VECTOR_SEMANTIC_HYBRID,
        top_k=5,
    )

    async with (
        AzureCliCredential() as credential,
        AzureAIAgentClient(
            project_endpoint=PROJECT_ENDPOINT,
            model_deployment_name=MODEL_DEPLOYMENT,
            async_credential=credential,
        ) as client,
    ):
        agent = client.create_agent(
            name=AGENT_NAME,
            instructions=INSTRUCTIONS,
            tools=ai_search.definitions,
            tool_resources=ai_search.resources,
        )
        agent_id: Optional[str] = getattr(agent, "id", None)
        print(f"[ok] agent created (id={agent_id or 'unknown'})\n")

        try:
            for i, q in enumerate(DEMO_QUESTIONS, 1):
                print(f"--- Q{i}: {q}")
                result = await agent.run(q)
                # AgentRunResponse stringifies to the assistant's final text.
                print(str(result).strip())
                print()
        finally:
            if not KEEP_AGENT and agent_id:
                try:
                    await client.project_client.agents.delete_agent(agent_id)
                    print(f"[cleanup] deleted agent {agent_id}")
                except Exception as e:  # noqa: BLE001
                    print(f"[cleanup] could not delete agent {agent_id}: {e}")

    return 0


def main() -> int:
    try:
        return asyncio.run(run())
    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":
    sys.exit(main())
