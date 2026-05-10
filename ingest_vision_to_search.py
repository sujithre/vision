#!/usr/bin/env python3
"""
Ingest the .vision.slides.jsonl produced by ``analyze_pptx_vision.py`` into an
Azure AI Search index with hybrid (BM25 + vector) + L2 semantic ranking.

All flat fields from the JSONL are first-class searchable fields:
  id, deck, slide_number, page_number,
  slide_title, slide_summary, slide_description,
  highlighted_terms, kpis_text, charts_text, tables_text, callouts_text,
  content (full slide markdown), and a 3072-dim `content_vector`.

Run AFTER ``analyze_pptx_vision.py``:

    python analyze_pptx_vision.py path/to/deck.pptx
    python ingest_vision_to_search.py path/to/deck.vision.slides.jsonl

Env (.env in cwd is auto-loaded):
    SEARCH_ENDPOINT          (default: https://copilotaisearchsuji.search.windows.net)
    SEARCH_INDEX             (default: chart-Automation)
    AOAI_ENDPOINT            (required - Foundry / AOAI endpoint hosting the embedding deployment)
    EMBED_DEPLOYMENT         (default: text-embedding-3-large-460208)
    EMBED_DIMS               (default: 3072)
    AOAI_API_VERSION         (default: 2024-10-21)

Auth: DefaultAzureCredential (az login). Required RBAC:
  - Search Service Contributor + Search Index Data Contributor on the Search service
  - Cognitive Services OpenAI User on the Foundry resource
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from azure.search.documents import SearchClient
from azure.search.documents.indexes import SearchIndexClient
from azure.search.documents.indexes.models import (
    HnswAlgorithmConfiguration,
    SearchableField,
    SearchField,
    SearchFieldDataType,
    SearchIndex,
    SemanticConfiguration,
    SemanticField,
    SemanticPrioritizedFields,
    SemanticSearch,
    SimpleField,
    VectorSearch,
    VectorSearchProfile,
)
from azure.search.documents.models import (
    QueryAnswerType,
    QueryCaptionType,
    QueryType,
    VectorizedQuery,
)
from dotenv import load_dotenv
from openai import AzureOpenAI

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

load_dotenv()

SEARCH_ENDPOINT = os.environ.get(
    "SEARCH_ENDPOINT", "https://copilotaisearchsuji.search.windows.net"
).rstrip("/")
INDEX_NAME = os.environ.get("SEARCH_INDEX", "chart-Automation")

AOAI_ENDPOINT = (
    os.environ.get("AOAI_ENDPOINT") or os.environ.get("FOUNDRY_ENDPOINT") or ""
).rstrip("/")
EMBED_DEPLOY = os.environ.get("EMBED_DEPLOYMENT", "text-embedding-3-large-460208")
EMBED_DIMS = int(os.environ.get("EMBED_DIMS", "3072"))
AOAI_API_VER = os.environ.get("AOAI_API_VERSION", "2024-10-21")

SEMANTIC_CONFIG = "slides-vision-semantic"
VECTOR_PROFILE = "slides-vision-vector-profile"
VECTOR_ALGO = "slides-vision-hnsw"

EMBED_FIELDS = [
    "slide_title",
    "slide_description",
    "slide_summary",
    "highlighted_terms",
    "kpis_text",
    "charts_text",
    "tables_text",
    "callouts_text",
    "content",
]


# ---------------------------------------------------------------------------
# Index definition
# ---------------------------------------------------------------------------

def build_index() -> SearchIndex:
    fields = [
        SimpleField(name="id", type=SearchFieldDataType.String, key=True, filterable=True),
        SimpleField(name="deck", type=SearchFieldDataType.String, filterable=True, facetable=True),
        SimpleField(name="slide_number", type=SearchFieldDataType.Int32, filterable=True, sortable=True),
        SimpleField(name="page_number", type=SearchFieldDataType.Int32, filterable=True, sortable=True),
        SearchableField(name="slide_title", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="slide_summary", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="slide_description", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="highlighted_terms", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="kpis_text", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="charts_text", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="tables_text", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="callouts_text", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchableField(name="content", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
        SearchField(
            name="content_vector",
            type=SearchFieldDataType.Collection(SearchFieldDataType.Single),
            searchable=True,
            vector_search_dimensions=EMBED_DIMS,
            vector_search_profile_name=VECTOR_PROFILE,
        ),
    ]
    vector_search = VectorSearch(
        algorithms=[HnswAlgorithmConfiguration(name=VECTOR_ALGO)],
        profiles=[
            VectorSearchProfile(
                name=VECTOR_PROFILE, algorithm_configuration_name=VECTOR_ALGO
            )
        ],
    )
    semantic = SemanticSearch(
        configurations=[
            SemanticConfiguration(
                name=SEMANTIC_CONFIG,
                prioritized_fields=SemanticPrioritizedFields(
                    title_field=SemanticField(field_name="slide_title"),
                    content_fields=[
                        SemanticField(field_name="slide_description"),
                        SemanticField(field_name="slide_summary"),
                        SemanticField(field_name="content"),
                    ],
                    keywords_fields=[
                        SemanticField(field_name="highlighted_terms"),
                        SemanticField(field_name="kpis_text"),
                        SemanticField(field_name="charts_text"),
                        SemanticField(field_name="tables_text"),
                        SemanticField(field_name="callouts_text"),
                    ],
                ),
            )
        ]
    )
    return SearchIndex(
        name=INDEX_NAME,
        fields=fields,
        vector_search=vector_search,
        semantic_search=semantic,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_jsonl(path: Path) -> list[dict]:
    docs: list[dict] = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                docs.append(json.loads(line))
    return docs


def embed_text_for(d: dict) -> str:
    parts = [str(d.get(f, "") or "").strip() for f in EMBED_FIELDS]
    return "\n\n".join(p for p in parts if p)


def embed_batch(aoai: AzureOpenAI, texts: list[str], batch: int = 16) -> list[list[float]]:
    out: list[list[float]] = []
    for i in range(0, len(texts), batch):
        chunk = texts[i : i + batch]
        resp = aoai.embeddings.create(model=EMBED_DEPLOY, input=chunk)
        out.extend(e.embedding for e in resp.data)
        print(f"    embedded {min(i + batch, len(texts))}/{len(texts)}")
    return out


def to_search_doc(d: dict, vector: list[float]) -> dict:
    return {
        "id": d["id"],
        "deck": d.get("deck", ""),
        "slide_number": int(d["slide_number"]),
        "page_number": int(d.get("page_number", d["slide_number"])),
        "slide_title": d.get("slide_title", "") or "",
        "slide_summary": d.get("slide_summary", "") or "",
        "slide_description": d.get("slide_description", "") or "",
        "highlighted_terms": d.get("highlighted_terms", "") or "",
        "kpis_text": d.get("kpis_text", "") or "",
        "charts_text": d.get("charts_text", "") or "",
        "tables_text": d.get("tables_text", "") or "",
        "callouts_text": d.get("callouts_text", "") or "",
        "content": d.get("content", "") or "",
        "content_vector": vector,
    }


# ---------------------------------------------------------------------------
# Optional sample query
# ---------------------------------------------------------------------------

def hybrid_semantic_search(
    aoai: AzureOpenAI, search_client: SearchClient, query: str, k: int = 5
) -> None:
    qvec = aoai.embeddings.create(model=EMBED_DEPLOY, input=query).data[0].embedding
    results = search_client.search(
        search_text=query,
        vector_queries=[
            VectorizedQuery(vector=qvec, k_nearest_neighbors=50, fields="content_vector")
        ],
        query_type=QueryType.SEMANTIC,
        semantic_configuration_name=SEMANTIC_CONFIG,
        query_caption=QueryCaptionType.EXTRACTIVE,
        query_answer=QueryAnswerType.EXTRACTIVE,
        select=[
            "id", "slide_number", "slide_title", "slide_summary",
            "highlighted_terms", "kpis_text", "charts_text",
            "tables_text", "callouts_text",
        ],
        top=k,
    )
    answers = results.get_answers() or []
    if answers:
        print("=== Semantic answers ===")
        for a in answers:
            print(f"  [{a.score:.2f}] {a.text}")
    print(f"\n=== Top {k} slides for: {query!r} ===")
    for r in results:
        rerank = r.get("@search.reranker_score")
        score = r["@search.score"]
        rerank_s = f"{rerank:.3f}" if rerank is not None else "n/a"
        print(f"\nslide {r['slide_number']:>2}  bm25/vec={score:.3f}  semantic={rerank_s}  - {r['slide_title']}")
        for label, key in [
            ("summary",    "slide_summary"),
            ("highlights", "highlighted_terms"),
            ("kpis",       "kpis_text"),
            ("charts",     "charts_text"),
            ("tables",     "tables_text"),
            ("callouts",   "callouts_text"),
        ]:
            v = r.get(key)
            if v:
                v = v if len(v) <= 300 else v[:300] + "..."
                print(f"  {label:<10}: {v}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("jsonl", type=Path,
                    help="Path to the *.vision.slides.jsonl produced by analyze_pptx_vision.py")
    ap.add_argument("--index", default=INDEX_NAME,
                    help=f"Search index name (default: {INDEX_NAME})")
    ap.add_argument("--skip-create", action="store_true",
                    help="Don't create_or_update the index (assume it already exists)")
    ap.add_argument("--query", default=None,
                    help="Optional sample query to run after upload")
    args = ap.parse_args()

    if not AOAI_ENDPOINT:
        print("ERROR: set AOAI_ENDPOINT (or FOUNDRY_ENDPOINT) in env / .env", file=sys.stderr)
        return 2
    if not args.jsonl.exists():
        print(f"ERROR: file not found: {args.jsonl}", file=sys.stderr)
        return 2

    global INDEX_NAME
    INDEX_NAME = args.index

    print(f"Search   : {SEARCH_ENDPOINT}")
    print(f"Index    : {INDEX_NAME}")
    print(f"Foundry  : {AOAI_ENDPOINT}")
    print(f"Embedder : {EMBED_DEPLOY}  (dims={EMBED_DIMS}, api={AOAI_API_VER})")
    print(f"Source   : {args.jsonl}\n")

    credential = DefaultAzureCredential()
    index_client = SearchIndexClient(endpoint=SEARCH_ENDPOINT, credential=credential)
    search_client = SearchClient(endpoint=SEARCH_ENDPOINT, index_name=INDEX_NAME, credential=credential)
    token_provider = get_bearer_token_provider(
        credential, "https://cognitiveservices.azure.com/.default"
    )
    aoai = AzureOpenAI(
        azure_endpoint=AOAI_ENDPOINT,
        azure_ad_token_provider=token_provider,
        api_version=AOAI_API_VER,
    )

    # Probe embedding dim.
    probe = aoai.embeddings.create(model=EMBED_DEPLOY, input="probe").data[0].embedding
    print(f"[1/4] embedding probe ok - dims={len(probe)} (expected {EMBED_DIMS})")
    if len(probe) != EMBED_DIMS:
        print(f"      WARNING: dim mismatch; index will be created with EMBED_DIMS={EMBED_DIMS}")

    # Index.
    if args.skip_create:
        print("[2/4] skip-create: assuming index exists")
    else:
        print("[2/4] creating/updating index...")
        result = index_client.create_or_update_index(build_index())
        print(f"      index '{result.name}' ready ({len(result.fields)} fields)")

    # Load + embed.
    docs = load_jsonl(args.jsonl)
    print(f"[3/4] loaded {len(docs)} slide docs; embedding...")
    vectors = embed_batch(aoai, [embed_text_for(d) for d in docs])

    # Upload.
    upload_docs = [to_search_doc(d, v) for d, v in zip(docs, vectors)]
    upload_result = search_client.upload_documents(documents=upload_docs)
    ok = sum(1 for r in upload_result if r.succeeded)
    print(f"[4/4] uploaded {ok}/{len(upload_result)} docs")
    for r in upload_result:
        if not r.succeeded:
            print(f"      FAILED {r.key}: {r.error_message}")

    if args.query:
        print()
        hybrid_semantic_search(aoai, search_client, args.query, k=5)

    return 0


if __name__ == "__main__":
    sys.exit(main())
