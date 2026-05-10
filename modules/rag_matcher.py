"""
rag_matcher.py — RAG-based classification layer for BS/PL labels.

Flow:
  1. Rule-based matcher (matcher.py) handles high-confidence labels → done.
  2. RAG retrieval: find top-K similar past labels from ChromaDB.
  3. If top-1 similarity ≥ RAG_CONFIDENCE_THRESHOLD → use its field directly.
  4. Otherwise call GPT-4o-mini with retrieved examples as in-context guidance.
  5. Every successful classification is stored for future retrieval.

The RAG DB is seeded on first run from matcher.py's BS_RULES and PL_RULES,
so even without any historical data, retrieval works from day 1.

Human corrections via /api/corrections get stored with confidence=1.0 and
outweigh any seeded rule-derived entries during retrieval.
"""
import os
import re
import hashlib
import logging
from pathlib import Path
from typing import Optional, List, Dict, Tuple
from datetime import datetime

# Lazy imports — these are heavy; only load if the module is used
_chroma = None
_st_model = None
_openai_client = None
_collection_bs = None
_collection_pl = None

_BASE_DIR = Path(__file__).parent
# Honor DATA_DIR-rooted paths from config.py so Railway can mount a volume
# at /data and persist ChromaDB + the embedding-model cache across deploys.
# Falls back to repo-root for local dev.
try:
    import sys as _sys
    _sys.path.insert(0, str(_BASE_DIR.parent))
    from config import CHROMA_DIR as _CHROMA_DIR, MODEL_CACHE_DIR as _MODEL_CACHE
except Exception:
    _CHROMA_DIR = _BASE_DIR.parent / "chroma_db"
    _MODEL_CACHE = _BASE_DIR.parent / "model_cache"

# Thresholds — tuned to be conservative (avoid false-confident mappings)
RAG_CONFIDENCE_THRESHOLD = 0.88    # top-1 similarity above this → use RAG directly
RAG_TOP_K = 5                       # how many examples to retrieve
GPT_MODEL = "gpt-4o-mini"
EMBED_MODEL_NAME = "sentence-transformers/all-MiniLM-L6-v2"

log = logging.getLogger("rag_matcher")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Lazy initialization
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _get_embedder():
    global _st_model
    if _st_model is None:
        from sentence_transformers import SentenceTransformer
        os.makedirs(_MODEL_CACHE, exist_ok=True)
        _st_model = SentenceTransformer(EMBED_MODEL_NAME,
                                        cache_folder=str(_MODEL_CACHE))
    return _st_model


def _get_collections():
    """Open (or create) the two ChromaDB collections: bs_labels, pl_labels."""
    global _chroma, _collection_bs, _collection_pl
    if _chroma is None:
        import chromadb
        os.makedirs(_CHROMA_DIR, exist_ok=True)
        _chroma = chromadb.PersistentClient(path=str(_CHROMA_DIR))
    if _collection_bs is None:
        _collection_bs = _chroma.get_or_create_collection(
            name="bs_labels",
            metadata={"hnsw:space": "cosine"},
        )
    if _collection_pl is None:
        _collection_pl = _chroma.get_or_create_collection(
            name="pl_labels",
            metadata={"hnsw:space": "cosine"},
        )
    return _collection_bs, _collection_pl


def _get_openai_client():
    """Read API key from config.json lazily on each call so config edits
    take effect without a process restart."""
    global _openai_client
    try:
        import sys
        sys.path.insert(0, str(_BASE_DIR.parent))
        from config import load_config
        api_key = load_config().get("openai_api_key", "").strip()
    except Exception:
        api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        return None
    try:
        from openai import OpenAI
        # Always create a fresh client if key changed
        _openai_client = OpenAI(api_key=api_key)
        return _openai_client
    except ImportError:
        return None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Seeding: populate DB from matcher.py rules on first run
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _label_id(label: str, field: str, section: str = "") -> str:
    """Stable ID from label+field+section, so the same entry isn't duplicated."""
    raw = f"{label}|{field}|{section}".lower()
    return hashlib.md5(raw.encode()).hexdigest()[:16]


def seed_from_rules():
    """Populate ChromaDB with representative examples derived from the
    rule patterns in matcher.py. Runs once at startup if collections are empty."""
    bs, pl = _get_collections()

    # Only seed if empty
    if bs.count() > 0 and pl.count() > 0:
        log.info(f"RAG DB already seeded (BS={bs.count()}, PL={pl.count()})")
        return

    # Seed data: mapping of sample labels → (field, section_hint)
    # These are hand-picked examples that represent each BS field. They give
    # the RAG retrieval something to match against from day 1.
    seed_bs = [
        ("Share capital", "Share Capital", "networth"),
        ("Equity share capital", "Share Capital", "networth"),
        ("Paid-up share capital", "Share Capital", "networth"),
        ("Retained earnings", "Retained Earnings", "networth"),
        ("Surplus in profit and loss", "Retained Earnings", "networth"),
        ("Reserves and surplus", "Reserves and Surplus", "networth"),
        ("General reserve", "Reserves and Surplus", "networth"),
        ("Securities premium", "Reserves and Surplus", "networth"),
        ("Capital reserve", "Reserves and Surplus", "networth"),
        ("Other equity", "Other Equity", "networth"),
        ("Share application money pending allotment", "Other Equity", "networth"),
        ("Money received against share warrants", "Other Equity", "networth"),
        ("Non-controlling interest", "Non Controlling Interest", "networth"),
        ("Minority interest", "Non Controlling Interest", "networth"),

        ("Trade payables", "Trade Payables", "current_liab"),
        ("Accounts payable", "Trade Payables", "current_liab"),
        ("Sundry creditors", "Trade Payables", "current_liab"),
        ("Total outstanding dues of micro enterprises and small enterprises",
         "Trade Payables", "current_liab"),
        ("Total outstanding dues of creditors other than micro enterprises",
         "Trade Payables", "current_liab"),
        ("Dues of micro and small enterprises", "Trade Payables", "current_liab"),

        ("Short-term borrowings", "Short Term Borrowings", "current_liab"),
        ("Bank overdraft", "Short Term Borrowings", "current_liab"),
        ("Cash credit", "Short Term Borrowings", "current_liab"),
        ("Current maturities of long-term debt", "Short Term Borrowings", "current_liab"),
        ("Long-term borrowings", "Long Term Borrowings", "noncurrent_liab"),
        ("Secured loans", "Long Term Borrowings", "noncurrent_liab"),
        ("Unsecured loans", "Long Term Borrowings", "noncurrent_liab"),
        ("Debentures", "Long Term Borrowings", "noncurrent_liab"),

        ("Short-term provisions", "Provisions Current", "current_liab"),
        ("Long-term provisions", "Provisions Non Current", "noncurrent_liab"),
        ("Employee benefit obligations", "Provisions Non Current", "noncurrent_liab"),
        ("Gratuity", "Provisions Non Current", "noncurrent_liab"),

        ("Lease liabilities (current)", "Lease Liabilities Current", "current_liab"),
        ("Current lease liabilities", "Lease Liabilities Current", "current_liab"),
        ("Lease liabilities (non-current)", "Other Non Current Liabilities", "noncurrent_liab"),

        ("Other current liabilities", "Other Current Liabilities", "current_liab"),
        ("Advance from customers", "Other Current Liabilities", "current_liab"),
        ("Income received in advance", "Other Current Liabilities", "current_liab"),
        ("Current tax liabilities (net)", "Other Current Liabilities", "current_liab"),
        ("Statutory dues payable", "Other Current Liabilities", "current_liab"),

        ("Other non-current liabilities", "Other Non Current Liabilities", "noncurrent_liab"),
        ("Deferred tax liabilities (net)", "Other Non Current Liabilities", "noncurrent_liab"),

        ("Other financial liabilities (current)", "Other Financial Liabilities Current", "current_liab"),
        ("Interest accrued but not due", "Other Financial Liabilities Current", "current_liab"),
        ("Unpaid dividend", "Other Financial Liabilities Current", "current_liab"),
        ("Other financial liabilities (non-current)",
         "Other Financial Liabilities Non Current", "noncurrent_liab"),

        ("Property, plant and equipment", "PPE Fixed Assets", "noncurrent_asset"),
        ("Tangible assets", "PPE Fixed Assets", "noncurrent_asset"),
        ("Intangible assets", "PPE Fixed Assets", "noncurrent_asset"),
        ("Goodwill", "PPE Fixed Assets", "noncurrent_asset"),
        ("Right-of-use assets", "PPE Fixed Assets", "noncurrent_asset"),
        ("Plant and machinery", "PPE Fixed Assets", "noncurrent_asset"),
        ("Investment property", "PPE Fixed Assets", "noncurrent_asset"),

        ("Capital work-in-progress", "CWIP", "noncurrent_asset"),
        ("Intangible assets under development", "CWIP", "noncurrent_asset"),

        ("Non-current investments", "Investments Non Current", "noncurrent_asset"),
        ("Investment in subsidiaries", "Investments Non Current", "noncurrent_asset"),
        ("Investment in associates", "Investments Non Current", "noncurrent_asset"),
        ("Current investments", "Current Investments", "current_asset"),

        ("Long-term loans and advances", "Loans Non Current", "noncurrent_asset"),
        ("Security deposits", "Loans Non Current", "noncurrent_asset"),
        ("Short-term loans and advances", "Loans Current", "current_asset"),

        ("Other non-current assets", "Other Non Current Assets", "noncurrent_asset"),
        ("Other financial assets (non-current)", "Other Financial Assets Non Current", "noncurrent_asset"),

        ("Deferred tax assets (net)", "Deferred Tax Assets", "noncurrent_asset"),

        ("Bank balances other than cash and cash equivalents", "Bank Balance", "current_asset"),
        ("Bank deposits", "Bank Balance", "current_asset"),
        ("Cash and cash equivalents", "Cash and Cash Equivalents", "current_asset"),
        ("Cash on hand", "Cash and Cash Equivalents", "current_asset"),

        ("Inventories", "Inventories", "current_asset"),
        ("Raw materials", "Inventories", "current_asset"),
        ("Finished goods", "Inventories", "current_asset"),
        ("Work in progress", "Inventories", "current_asset"),

        ("Trade receivables", "Trade Receivables", "current_asset"),
        ("Accounts receivable", "Trade Receivables", "current_asset"),
        ("Sundry debtors", "Trade Receivables", "current_asset"),

        ("Other current assets", "Other Current Assets", "current_asset"),
        ("Current tax assets (net)", "Other Current Assets", "current_asset"),
        ("Prepaid expenses", "Other Current Assets", "current_asset"),
        ("Other financial assets (current)", "Other Financial Assets Current", "current_asset"),
    ]

    seed_pl = [
        ("Revenue from operations", "Revenue", ""),
        ("Net sales", "Revenue", ""),
        ("Sale of products", "Revenue", ""),
        ("Sale of services", "Revenue", ""),
        ("Operating revenue", "Revenue", ""),

        ("Cost of materials consumed", "Cost of goods sold", ""),
        ("Purchase of stock-in-trade", "Cost of goods sold", ""),
        ("Changes in inventories of finished goods", "Cost of goods sold", ""),

        ("Employee benefits expense", "Employee benefits expense", ""),
        ("Salaries and wages", "Employee benefits expense", ""),
        ("Staff welfare expenses", "Employee benefits expense", ""),

        ("Finance costs", "Interest", ""),
        ("Interest expense", "Interest", ""),
        ("Interest on borrowings", "Interest", ""),

        ("Depreciation and amortisation expense", "Depreciation", ""),
        ("Depreciation", "Depreciation", ""),
        ("Amortisation of intangibles", "Depreciation", ""),

        ("Current tax", "Taxes", ""),
        ("Deferred tax", "Taxes", ""),
        ("Tax expense", "Taxes", ""),
        ("Income tax expense", "Taxes", ""),
        ("Current tax (net of MAT credit)", "Taxes", ""),

        ("Other income", "other_income", ""),
        ("Interest income", "other_income", ""),
        ("Dividend income", "other_income", ""),

        ("Other expenses", "Other Expenses less Other Income", ""),
        ("Administrative expenses", "Other Expenses less Other Income", ""),
        ("Selling expenses", "Other Expenses less Other Income", ""),

        ("Exceptional items", "exceptional", ""),
        ("Profit on sale of investments (exceptional)", "exceptional", ""),
    ]

    embedder = _get_embedder()

    # Bulk seed BS
    if bs.count() == 0:
        docs = [s[0] for s in seed_bs]
        embeds = embedder.encode(docs, convert_to_numpy=True).tolist()
        ids = [_label_id(s[0], s[1], s[2]) for s in seed_bs]
        metas = [{"field": s[1], "section": s[2], "source": "seed",
                  "confidence": 0.80, "added_at": datetime.now().isoformat()}
                 for s in seed_bs]
        bs.add(ids=ids, documents=docs, embeddings=embeds, metadatas=metas)
        log.info(f"Seeded {len(seed_bs)} BS entries")

    # Bulk seed PL
    if pl.count() == 0:
        docs = [s[0] for s in seed_pl]
        embeds = embedder.encode(docs, convert_to_numpy=True).tolist()
        ids = [_label_id(s[0], s[1]) for s in seed_pl]
        metas = [{"field": s[1], "source": "seed", "confidence": 0.80,
                  "added_at": datetime.now().isoformat()}
                 for s in seed_pl]
        pl.add(ids=ids, documents=docs, embeddings=embeds, metadatas=metas)
        log.info(f"Seeded {len(seed_pl)} PL entries")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Retrieval
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _retrieve_bs(label: str, section: Optional[str] = None,
                 top_k: int = RAG_TOP_K) -> List[Dict]:
    """Retrieve top-K most similar past BS labels. Optionally filter by section."""
    bs, _ = _get_collections()
    if bs.count() == 0:
        return []
    embedder = _get_embedder()
    query_emb = embedder.encode([label], convert_to_numpy=True).tolist()

    where = None
    if section:
        where = {"section": section}
    try:
        results = bs.query(
            query_embeddings=query_emb,
            n_results=min(top_k, bs.count()),
            where=where,
        )
    except Exception:
        # Section filter too strict → try unfiltered
        results = bs.query(query_embeddings=query_emb,
                           n_results=min(top_k, bs.count()))
    return _parse_results(results)


def _retrieve_pl(label: str, top_k: int = RAG_TOP_K) -> List[Dict]:
    _, pl = _get_collections()
    if pl.count() == 0:
        return []
    embedder = _get_embedder()
    query_emb = embedder.encode([label], convert_to_numpy=True).tolist()
    results = pl.query(query_embeddings=query_emb,
                       n_results=min(top_k, pl.count()))
    return _parse_results(results)


def _parse_results(results: dict) -> List[Dict]:
    """ChromaDB returns columnar data; convert to a list of row dicts with
    similarity score (1 - cosine distance)."""
    if not results or not results.get("ids") or not results["ids"][0]:
        return []
    out = []
    for i in range(len(results["ids"][0])):
        distance = results["distances"][0][i] if results.get("distances") else 1.0
        similarity = max(0.0, 1.0 - distance)
        out.append({
            "label": results["documents"][0][i],
            "field": results["metadatas"][0][i].get("field"),
            "section": results["metadatas"][0][i].get("section", ""),
            "similarity": similarity,
            "confidence": results["metadatas"][0][i].get("confidence", 0.5),
            "source": results["metadatas"][0][i].get("source", "unknown"),
        })
    return out


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# GPT-4o-mini fallback
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_FIELD_LIST = [
    "Share Capital", "Retained Earnings", "Reserves and Surplus", "Other Equity",
    "Non Controlling Interest",
    "Trade Payables", "Provisions Current", "Short Term Borrowings",
    "Other Current Liabilities", "Other Financial Liabilities Current",
    "Lease Liabilities Current",
    "Long Term Borrowings", "Provisions Non Current", "Other Non Current Liabilities",
    "Other Financial Liabilities Non Current",
    "PPE Fixed Assets", "Investments Non Current", "Loans Non Current", "CWIP",
    "Other Non Current Assets", "Other Financial Assets Non Current",
    "Deferred Tax Assets",
    "Bank Balance", "Cash and Cash Equivalents", "Inventories", "Current Investments",
    "Loans Current", "Trade Receivables", "Other Current Assets",
    "Other Financial Assets Current",
    "_skip",
]

PL_FIELD_LIST = [
    "Revenue", "Cost of goods sold", "Employee benefits expense", "Interest",
    "Depreciation", "Taxes", "other_income", "exceptional",
    "Other Expenses less Other Income", "_skip",
]


def _gpt_classify(label: str, field_list: list, examples: List[Dict],
                  side: str = "BS", section: str = "") -> Tuple[Optional[str], float, str]:
    """Ask GPT-4o-mini to classify a label into one of field_list.
    Returns (field, confidence, reasoning)."""
    client = _get_openai_client()
    if client is None:
        return None, 0.0, "no_api_key"

    # Build a tight prompt with retrieved examples
    examples_text = "\n".join(
        f"  - \"{e['label']}\" → {e['field']}  (similarity: {e['similarity']:.2f})"
        for e in examples[:5]
    ) or "  (no similar examples available)"

    field_list_str = "\n".join(f"  - {f}" for f in field_list)

    section_hint = f" The label appears in the {section} section." if section else ""

    prompt = f"""You are classifying a line item from an Indian annual report's {side}.{section_hint}

Label to classify:
  "{label}"

Similar past labels and their correct fields:
{examples_text}

Allowed fields:
{field_list_str}

Respond with ONLY a JSON object:
{{"field": "<one of the allowed fields>", "confidence": <0.0 to 1.0>, "reason": "<brief>"}}

If the label is a section header or total row that should not be mapped, return field="_skip".
"""
    try:
        resp = client.chat.completions.create(
            model=GPT_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=150,
            response_format={"type": "json_object"},
        )
        import json
        payload = json.loads(resp.choices[0].message.content)
        field = payload.get("field", "").strip()
        confidence = float(payload.get("confidence", 0.5))
        reason = payload.get("reason", "")[:120]
        if field in field_list:
            return field, confidence, reason
        return None, 0.0, f"gpt_returned_invalid_field: {field}"
    except Exception as e:
        return None, 0.0, f"gpt_error: {str(e)[:80]}"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public classification API
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def classify_bs_label(label: str, section: Optional[str] = None,
                      use_gpt: bool = True) -> Dict:
    """
    Classify a BS label using RAG + optional GPT-4o-mini fallback.

    Returns:
      {
        "field": "Trade Payables" | None,
        "confidence": 0.0..1.0,
        "source": "rag" | "gpt" | "none",
        "retrieved": [...top-5 examples...],
        "reasoning": "...",
      }
    """
    examples = _retrieve_bs(label, section=section, top_k=RAG_TOP_K)

    # Direct RAG hit
    if examples and examples[0]["similarity"] >= RAG_CONFIDENCE_THRESHOLD:
        top = examples[0]
        return {
            "field": top["field"],
            "confidence": top["similarity"] * top["confidence"],
            "source": "rag",
            "retrieved": examples,
            "reasoning": f"RAG direct hit (sim={top['similarity']:.3f}): {top['label']}",
        }

    # GPT fallback
    if use_gpt:
        field, conf, reason = _gpt_classify(label, BS_FIELD_LIST, examples,
                                            side="BS", section=section or "")
        if field:
            return {
                "field": field,
                "confidence": conf,
                "source": "gpt",
                "retrieved": examples,
                "reasoning": reason,
            }

    # Best-effort: return the top retrieved match if any
    if examples:
        top = examples[0]
        return {
            "field": top["field"],
            "confidence": top["similarity"] * 0.5,  # penalized
            "source": "rag_low",
            "retrieved": examples,
            "reasoning": f"low-confidence RAG (sim={top['similarity']:.3f})",
        }

    return {"field": None, "confidence": 0.0, "source": "none",
            "retrieved": [], "reasoning": "no match and no GPT available"}


def classify_pl_label(label: str, use_gpt: bool = True) -> Dict:
    examples = _retrieve_pl(label, top_k=RAG_TOP_K)
    if examples and examples[0]["similarity"] >= RAG_CONFIDENCE_THRESHOLD:
        top = examples[0]
        return {
            "field": top["field"],
            "confidence": top["similarity"] * top["confidence"],
            "source": "rag",
            "retrieved": examples,
            "reasoning": f"RAG direct hit (sim={top['similarity']:.3f}): {top['label']}",
        }
    if use_gpt:
        field, conf, reason = _gpt_classify(label, PL_FIELD_LIST, examples, side="PL")
        if field:
            return {
                "field": field, "confidence": conf, "source": "gpt",
                "retrieved": examples, "reasoning": reason,
            }
    if examples:
        top = examples[0]
        return {
            "field": top["field"], "confidence": top["similarity"] * 0.5,
            "source": "rag_low", "retrieved": examples,
            "reasoning": f"low-confidence RAG (sim={top['similarity']:.3f})",
        }
    return {"field": None, "confidence": 0.0, "source": "none",
            "retrieved": [], "reasoning": "no match"}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Learning: add new verified examples
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def add_correction(label: str, field: str, side: str = "BS",
                   section: str = "", source_company: str = "") -> bool:
    """Called when a human reviewer confirms the correct mapping for a label.
    This entry outweighs seed data in future retrieval (confidence=1.0)."""
    bs, pl = _get_collections()
    embedder = _get_embedder()
    collection = bs if side.upper() == "BS" else pl

    emb = embedder.encode([label], convert_to_numpy=True).tolist()
    id_ = _label_id(label, field, section)
    meta = {
        "field": field,
        "section": section,
        "source": f"correction:{source_company}" if source_company else "correction",
        "confidence": 1.0,
        "added_at": datetime.now().isoformat(),
    }
    try:
        # Upsert: either adds new or updates existing with higher confidence
        collection.upsert(ids=[id_], documents=[label],
                          embeddings=emb, metadatas=[meta])
        log.info(f"Correction added: [{side}] '{label}' → {field}")
        return True
    except Exception as e:
        log.warning(f"Failed to add correction: {e}")
        return False


def stats() -> Dict:
    """Return DB sizes and source breakdowns."""
    bs, pl = _get_collections()
    return {
        "bs_count": bs.count(),
        "pl_count": pl.count(),
        "chroma_path": str(_CHROMA_DIR),
        "embed_model": EMBED_MODEL_NAME,
        "gpt_model": GPT_MODEL,
        "rag_threshold": RAG_CONFIDENCE_THRESHOLD,
    }


# On module import, attempt to seed. Fails silently if chromadb not installed.
try:
    seed_from_rules()
except Exception as e:
    log.warning(f"RAG DB seeding skipped: {e}")
