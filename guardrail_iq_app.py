"""
GuardrailIQ — Deterministic GRC Compliance Reviewer (single-file Streamlit app)
================================================================================

Turns a governance Excel workbook (Governance_Documents, Clauses,
Design_Descriptions) into a structured, traceable, de-duplicated guardrail model,
then evaluates uploaded design documents (DOCX / PDF / PPTX / TXT) against it.

Design principles (per the GuardrailIQ challenge):
  * NO external LLM. Everything is deterministic, rules + classical NLP, so every
    verdict is fully explainable and reproducible.
  * Every finding cites BOTH the guardrail AND its source clause(s). No verdict is
    ever emitted without a traceable citation.
  * Detection is by MEANING (keywords / numeric parsing / graph joins), never by
    memorised record IDs, so it survives a "modified copy" of the dataset.

Pipeline
  Excel --(flexible alias mapping)--> normalised tables
        --> severity-tagged guardrails (one per clause)
        --> TF-IDF + cosine + NetworkX connected-components de-duplication
            (all source clauses preserved on merge)
        --> numeric-threshold conflict detection (surfaced, never auto-resolved)
        --> design ingestion (DOCX/PDF/PPTX/TXT) + keyword-evidence evaluation
            (Compliant / Non-compliant / Gap, each with clause citations)
        --> dashboard, remediation planner, multi-sheet audit export.

Run:  streamlit run guardrail_iq_app.py
"""

from __future__ import annotations

import io
import re
import json
from datetime import datetime
from collections import defaultdict, Counter

import numpy as np
import pandas as pd
import streamlit as st
import networkx as nx
import plotly.express as px
import plotly.graph_objects as go
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ----------------------------------------------------------------------------- #
# 0. Page config & light styling
# ----------------------------------------------------------------------------- #
st.set_page_config(page_title="GuardrailIQ — GRC Compliance Reviewer",
                   page_icon="🛡️", layout="wide")

SEVERITY_ORDER = {"High": 3, "Medium": 2, "Low": 1}
SEVERITY_COLOR = {"High": "#C0392B", "Medium": "#E67E22", "Low": "#2E86C1"}
VERDICT_COLOR = {"Compliant": "#27AE60", "Non-compliant": "#C0392B",
                 "Gap": "#F1C40F", "Not applicable": "#95A5A6"}


# ============================================================================= #
# 1. FLEXIBLE EXCEL INGESTION (schema / alias mapping)
# ============================================================================= #
# Canonical column -> set of accepted aliases (normalised: lowercase, alnum only).
SHEET_SPECS = {
    "governance_documents": {
        "canonical_id": "Document_ID",
        "aliases": {
            "Document_ID": ["documentid", "docid", "id", "documentidentifier"],
            "Document_Title": ["documenttitle", "title", "name", "documentname"],
            "Domain": ["domain", "governancedomain", "area"],
            "Document_Type": ["documenttype", "type", "doctype"],
            "Owner_Function": ["ownerfunction", "owner", "function", "responsible"],
            "Version": ["version", "ver"],
            "Effective_Date": ["effectivedate", "date", "effective"],
            "Status": ["status", "state"],
            "Summary": ["summary", "description", "abstract"],
        },
    },
    "clauses": {
        "canonical_id": "Clause_ID",
        "aliases": {
            "Clause_ID": ["clauseid", "id", "clauseidentifier"],
            "Document_ID": ["documentid", "docid", "parentdocument", "parentid"],
            "Clause_Ref": ["clauseref", "ref", "reference", "section", "clausenumber"],
            "Clause_Heading": ["clauseheading", "heading", "title", "clausetitle"],
            "Clause_Text": ["clausetext", "text", "body", "content", "requirement"],
            "Seeded_Flag": ["seededflag", "flag", "seed", "note"],
        },
    },
    "design_descriptions": {
        "canonical_id": "Design_ID",
        "aliases": {
            "Design_ID": ["designid", "id", "designidentifier"],
            "Design_Title": ["designtitle", "title", "name"],
            "Domain": ["domain", "area"],
            "Design_Description": ["designdescription", "description", "text", "body"],
            "Expected_Outcome": ["expectedoutcome", "expected", "outcome", "label"],
            "Relevant_Guardrails": ["relevantguardrails", "guardrails"],
            "Scenario_Note": ["scenarionote", "note", "notes"],
        },
    },
}


def _norm(s: str) -> str:
    """Normalise a header/sheet name to lowercase alphanumerics only."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def match_sheet(sheet_names, target_key):
    """Fuzzily match a workbook sheet name to a canonical sheet key."""
    tnorm = _norm(target_key)
    # exact-ish match first
    for s in sheet_names:
        if _norm(s) == tnorm:
            return s
    # token-overlap fallback
    tkey = target_key.split("_")[0]  # e.g. 'governance', 'clauses', 'design'
    for s in sheet_names:
        if _norm(tkey) in _norm(s):
            return s
    return None


def map_columns(df, spec):
    """Rename df columns to canonical names using alias sets. Returns new df."""
    alias_lookup = {}
    for canonical, aliases in spec["aliases"].items():
        alias_lookup[_norm(canonical)] = canonical
        for a in aliases:
            alias_lookup[_norm(a)] = canonical
    rename = {}
    for col in df.columns:
        key = _norm(col)
        if key in alias_lookup:
            rename[col] = alias_lookup[key]
    out = df.rename(columns=rename)
    # keep only canonical columns that exist; ensure all canonical exist
    for canonical in spec["aliases"]:
        if canonical not in out.columns:
            out[canonical] = ""
    out = out[list(spec["aliases"].keys())]
    return out


def load_workbook(file_bytes):
    """Read the workbook and return (docs_df, clauses_df, designs_df, report)."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    report = {"sheets_found": xls.sheet_names, "mapping": {}}
    frames = {}
    for key, spec in SHEET_SPECS.items():
        sheet = match_sheet(xls.sheet_names, key)
        if sheet is None:
            frames[key] = pd.DataFrame(columns=list(spec["aliases"].keys()))
            report["mapping"][key] = {"matched_sheet": None, "columns": {}}
            continue
        raw = xls.parse(sheet, dtype=str).fillna("")
        mapped = map_columns(raw, spec)
        frames[key] = mapped
        report["mapping"][key] = {
            "matched_sheet": sheet,
            "columns": {c: (c in [ _canon_from_raw(rc, spec) for rc in raw.columns ])
                        for c in spec["aliases"]},
            "raw_columns": list(raw.columns),
        }
    return (frames["governance_documents"], frames["clauses"],
            frames["design_descriptions"], report)


def _canon_from_raw(raw_col, spec):
    key = _norm(raw_col)
    for canonical, aliases in spec["aliases"].items():
        if key == _norm(canonical) or key in [_norm(a) for a in aliases]:
            return canonical
    return None


# ============================================================================= #
# 2. SEVERITY TAGGING (keyword-weighted, deterministic)
# ============================================================================= #
SEVERITY_RULES = [
    ("High", ["mfa", "multi-factor", "multifactor", "encrypt", "pii", "personal data",
              "public", "anonymous", "least privilege", "least-privilege", "admin",
              "administrative", "penetration", "critical", "residency", "subject",
              "erasure", "secret", "hardcoded", "credential", "restricted",
              "privileged", "revoke", "internet-facing", "public internet",
              "classif", "sast", "dast"]),
    ("Medium", ["log", "logging", "monitor", "review", "recertif", "backup", "restore",
                "patch", "vulnerab", "segment", "firewall", "rotation", "rotate",
                "code review", "peer review", "dependency", "retention", "retain",
                "purge", "assessment", "rto", "rpo", "supplier", "third party",
                "third-party", "risk assessment", "password"]),
    ("Low", ["principle", "guidance", "recommended", "consider", "should", "prototype"]),
]


def tag_severity(text: str) -> str:
    """Assign High/Medium/Low by counting weighted keyword hits (deterministic)."""
    t = " " + text.lower() + " "
    scores = {"High": 0, "Medium": 0, "Low": 0}
    for sev, kws in SEVERITY_RULES:
        for kw in kws:
            if kw in t:
                scores[sev] += 1
    # High dominates if present; else Medium; else Low; default Medium
    if scores["High"] > 0:
        return "High"
    if scores["Medium"] > 0:
        return "Medium"
    if scores["Low"] > 0:
        return "Low"
    return "Medium"


CATEGORY_RULES = [
    ("Identity & Access Management", ["mfa", "multi-factor", "least privilege", "password",
                                      "access", "privileged", "recertif", "shared account",
                                      "joiner", "leaver", "revoke", "authentication"]),
    ("Encryption & Key Management", ["encrypt", "tls", "aes", "key", "cipher", "cryptograph"]),
    ("Data Protection & Privacy", ["pii", "personal data", "residency", "subject",
                                   "retention", "purge", "anonymise", "anonymize", "dlp",
                                   "classif"]),
    ("Logging & Monitoring", ["log", "monitor", "alert", "audit event", "siem"]),
    ("Application Security", ["sast", "dast", "code review", "peer review", "dependency",
                              "secret", "hardcoded", "pipeline", "ci/cd"]),
    ("Network Security", ["firewall", "segment", "port", "admin interface", "perimeter",
                          "internet-facing", "public internet"]),
    ("Cloud & Infrastructure", ["cloud", "bucket", "object store", "well-architected",
                                "infrastructure as code", "public access"]),
    ("Resilience", ["backup", "restore", "rto", "rpo", "continuity", "recover"]),
    ("Vulnerability Management", ["vulnerab", "patch", "remediat", "penetration", "pen test"]),
    ("Third-Party Security", ["supplier", "third party", "third-party", "vendor"]),
]


def tag_category(text: str) -> str:
    t = text.lower()
    for cat, kws in CATEGORY_RULES:
        if any(k in t for k in kws):
            return cat
    return "General Governance"


# ============================================================================= #
# 3. BUILD GUARDRAILS FROM CLAUSES
# ============================================================================= #
def build_guardrails(clauses_df, docs_df):
    """One raw guardrail per clause, enriched with domain/category/severity."""
    doc_domain = dict(zip(docs_df["Document_ID"], docs_df["Domain"]))
    doc_title = dict(zip(docs_df["Document_ID"], docs_df["Document_Title"]))
    doc_ids = set(docs_df["Document_ID"])

    rows = []
    for _, c in clauses_df.iterrows():
        cid = str(c["Clause_ID"]).strip()
        if not cid:
            continue
        did = str(c["Document_ID"]).strip()
        heading = str(c["Clause_Heading"]).strip()
        text = str(c["Clause_Text"]).strip()
        blob = f"{heading}. {text}".strip(". ")
        domain = doc_domain.get(did, "") or tag_category(blob)
        rows.append({
            "guardrail_id": f"GR-{cid}",
            "statement": heading if heading else (text[:80] + "…"),
            "requirement": text,
            "domain": domain,
            "category": tag_category(blob),
            "severity": tag_severity(blob),
            "source_clause_ids": [cid],
            "source_document_ids": [did] if did else [],
            "clause_ref": str(c["Clause_Ref"]).strip(),
            "seeded_flag": str(c["Seeded_Flag"]).strip(),
            "_text": blob,
            "orphan_refs": [] if did in doc_ids or not did else [did],
        })
    gdf = pd.DataFrame(rows)
    return gdf


# ============================================================================= #
# 4. DE-DUPLICATION  (TF-IDF + cosine + NetworkX connected components)
# ============================================================================= #
def _measure_signature(text):
    """Return {measure_name: normalised_value} for numeric measures in text."""
    return {name: norm for name, raw, norm in _extract_measures(text) if norm is not None}


def deduplicate(gdf, sim_threshold=0.26):
    """
    Cluster near-duplicate guardrails by semantic similarity of their text
    (TF-IDF + cosine), grouped via NetworkX connected components. Merge each
    cluster into ONE canonical guardrail, preserving ALL source clause/document
    references.

    Safety rule: two guardrails that state DIFFERENT numeric values for the SAME
    measure are a *conflict*, not a duplicate — such edges are never created,
    so conflicting requirements (e.g. 8- vs 12-char passwords) are never merged.

    Returns (deduped_df, clusters, sim_matrix, ids).
    """
    texts = gdf["_text"].tolist()
    ids = gdf["guardrail_id"].tolist()
    if len(texts) < 2:
        gdf = gdf.copy()
        gdf["merged_from"] = [[i] for i in ids]
        gdf["is_merged"] = False
        return gdf, [], np.zeros((len(texts), len(texts))), ids

    vec = TfidfVectorizer(stop_words="english", ngram_range=(1, 2), min_df=1)
    tfidf = vec.fit_transform(texts)
    sim = cosine_similarity(tfidf)

    signatures = [_measure_signature(t) for t in texts]

    def _conflicting(i, j):
        for m, v in signatures[i].items():
            if m in signatures[j] and abs(signatures[j][m] - v) > 1e-6:
                return True
        return False

    G = nx.Graph()
    G.add_nodes_from(range(len(ids)))
    for i in range(len(ids)):
        for j in range(i + 1, len(ids)):
            if sim[i, j] >= sim_threshold and not _conflicting(i, j):
                G.add_edge(i, j, weight=round(float(sim[i, j]), 3))

    clusters = [sorted(c) for c in nx.connected_components(G)]

    merged_rows = []
    cluster_info = []
    for comp in clusters:
        members = gdf.iloc[comp]
        # canonical = highest severity, then longest requirement text
        members_sorted = members.assign(
            _sev=members["severity"].map(SEVERITY_ORDER),
            _len=members["requirement"].str.len(),
        ).sort_values(["_sev", "_len"], ascending=False)
        canon = members_sorted.iloc[0].to_dict()

        all_clauses, all_docs, all_gids, all_orphans = [], [], [], []
        for _, m in members.iterrows():
            all_clauses += m["source_clause_ids"]
            all_docs += m["source_document_ids"]
            all_gids.append(m["guardrail_id"])
            all_orphans += m["orphan_refs"]
        canon["source_clause_ids"] = sorted(set(all_clauses))
        canon["source_document_ids"] = sorted(set(all_docs))
        canon["orphan_refs"] = sorted(set(all_orphans))
        canon["merged_from"] = sorted(set(all_gids))
        canon["is_merged"] = len(comp) > 1
        canon["severity"] = max(members["severity"], key=lambda s: SEVERITY_ORDER[s])
        merged_rows.append(canon)

        if len(comp) > 1:
            cluster_info.append({
                "canonical": canon["guardrail_id"],
                "members": all_gids,
                "clauses": canon["source_clause_ids"],
                "avg_similarity": round(float(np.mean(
                    [sim[i, j] for i in comp for j in comp if i < j])), 3),
                "statement": canon["statement"],
            })

    ddf = pd.DataFrame(merged_rows).drop(columns=["_sev", "_len"], errors="ignore")
    ddf = ddf.reset_index(drop=True)
    return ddf, cluster_info, sim, ids


# ============================================================================= #
# 5. NUMERIC-THRESHOLD CONFLICT DETECTION (surfaced, never auto-resolved)
# ============================================================================= #
# Each measure: name -> (trigger keywords, regex to extract a numeric value+unit)
MEASURE_PATTERNS = [
    ("Password minimum length",
     ["password", "characters", "character"],
     r"(\d+)\s*[- ]?character"),
    ("Log / data retention period",
     ["log", "retention", "retain", "purge", "anonymise", "anonymize"],
     r"(\d+)\s*(day|days|month|months|year|years)"),
    ("Cryptographic key rotation period",
     ["key", "rotat"],
     r"(\d+)\s*(day|days|month|months|year|years)"),
    ("Access recertification frequency",
     ["recertif", "access review", "review access"],
     r"(\d+)\s*(day|days|month|months|quarter|quarters)|\b(quarter|quarterly|annual|annually|monthly)\b"),
    ("Critical vulnerability remediation SLA",
     ["vulnerab", "remediat", "patch"],
     r"(\d+)\s*(day|days|hour|hours)"),
    ("Incident reporting window",
     ["incident", "report"],
     r"(\d+)\s*(hour|hours|day|days)"),
    ("Access revocation window",
     ["revoke", "leaver", "joiner", "leaving"],
     r"(\d+)\s*(hour|hours|day|days)"),
    ("Backup frequency",
     ["backup", "back up"],
     r"\b(daily|weekly|monthly|hourly)\b"),
    ("TLS minimum version",
     ["tls"],
     r"tls\s*(\d\.\d)"),
]

_UNIT_TO_DAYS = {"hour": 1/24, "hours": 1/24, "day": 1, "days": 1,
                 "month": 30, "months": 30, "year": 365, "years": 365,
                 "quarter": 90, "quarters": 90}
_FREQ_TO_DAYS = {"hourly": 1/24, "daily": 1, "weekly": 7, "monthly": 30,
                 "quarterly": 90, "quarter": 90, "annual": 365, "annually": 365}


def _extract_measures(text):
    """Return list of (measure_name, raw_value_str, normalised_numeric)."""
    t = text.lower()
    found = []
    for name, triggers, pattern in MEASURE_PATTERNS:
        if not any(k in t for k in triggers):
            continue
        m = re.search(pattern, t)
        if not m:
            continue
        raw = m.group(0).strip()
        norm = None
        groups = [g for g in m.groups() if g]
        # numeric + unit
        num = next((g for g in groups if g and g.isdigit()), None)
        unit = next((g for g in groups if g and not g.isdigit()), None)
        if num is not None and unit in _UNIT_TO_DAYS:
            norm = float(num) * _UNIT_TO_DAYS[unit]
        elif num is not None and name == "Password minimum length":
            norm = float(num)          # chars, keep as-is
        elif num is not None and name == "TLS minimum version":
            norm = float(num)
        elif num is not None:
            norm = float(num)
        elif unit in _FREQ_TO_DAYS:
            norm = _FREQ_TO_DAYS[unit]
        elif raw in _FREQ_TO_DAYS:
            norm = _FREQ_TO_DAYS[raw]
        found.append((name, raw, norm))
    return found


def detect_conflicts(ddf):
    """
    Group guardrails by measure and surface DIFFERENT numeric values for the same
    measure as a conflict. Does NOT resolve — presents all sides with citations.
    """
    buckets = defaultdict(list)  # measure_name -> list of dicts
    for _, g in ddf.iterrows():
        for name, raw, norm in _extract_measures(g["_text"]):
            if norm is None:
                continue
            buckets[name].append({
                "guardrail_id": g["guardrail_id"],
                "statement": g["statement"],
                "value_raw": raw,
                "value_norm": norm,
                "clauses": g["source_clause_ids"],
                "documents": g["source_document_ids"],
                "domain": g["domain"],
            })

    conflicts = []
    for measure, entries in buckets.items():
        distinct = sorted(set(round(e["value_norm"], 4) for e in entries))
        if len(distinct) > 1:
            conflicts.append({
                "measure": measure,
                "distinct_values": distinct,
                "positions": entries,
                "resolution": "UNRESOLVED — human decision required",
            })
    return conflicts


# ============================================================================= #
# 6. DESIGN DOCUMENT INGESTION (DOCX / PDF / PPTX / TXT)
# ============================================================================= #
def extract_text(uploaded):
    """Return plain text from an uploaded design file, by extension."""
    name = uploaded.name.lower()
    data = uploaded.read()
    if name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        import docx
        d = docx.Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for tbl in d.tables:
            for row in tbl.rows:
                parts.append(" ".join(cell.text for cell in row.cells))
        return "\n".join(parts)
    if name.endswith(".pdf"):
        try:
            import fitz
            doc = fitz.open(stream=data, filetype="pdf")
            return "\n".join(page.get_text() for page in doc)
        except Exception:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                return "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    if name.endswith(".pptx"):
        from pptx import Presentation
        prs = Presentation(io.BytesIO(data))
        parts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    parts.append(shape.text_frame.text)
        return "\n".join(parts)
    return data.decode("utf-8", errors="ignore")


# ============================================================================= #
# 7. EVIDENCE-BASED EVALUATION ENGINE (deterministic detectors)
# ============================================================================= #
# Each detector encodes MEANING, not record IDs:
#   trigger   -> topic relevance (is this control in scope for the design?)
#   comply    -> phrases showing the control is satisfied
#   violate   -> phrases showing the control is broken
# A guardrail is linked to a detector if the detector trigger keywords appear in
# the guardrail statement/requirement. Evaluation then runs the linked detectors
# against the design text.
DETECTORS = [
    {"key": "public_storage",
     "trigger": ["public", "bucket", "object store", "anonymous", "storage"],
     "gr_link": ["public", "bucket", "object store", "anonymous", "storage"],
     "comply": ["public access is blocked", "public access block", "no public",
                "not publicly", "public-access-block", "access blocked",
                "storage public access is blocked"],
     "violate": ["public anonymous read", "publicly accessible", "public read",
                 "world-readable", "anonymous read", "public anonymous",
                 "configured for public", "public/anonymous"]},
    {"key": "mfa",
     "trigger": ["mfa", "multi-factor", "multifactor"],
     "gr_link": ["mfa", "multi-factor", "multifactor"],
     "comply": ["mfa for admins", "mfa enforced", "with mfa", "mfa enabled",
                "multi-factor", "requires mfa", "sso with mfa"],
     "violate": ["no mfa", "single-factor", "without mfa", "only by a username",
                 "only a username and password", "username and password. no mfa"]},
    {"key": "encryption_at_rest",
     "trigger": ["encrypt", "at rest", "aes"],
     "gr_link": ["encrypt", "at rest", "aes"],
     "comply": ["encrypted at rest", "encryption at rest", "aes-256",
                "data encrypted at rest", "storage encryption"],
     "violate": ["no encryption at rest", "unencrypted", "not encrypted",
                 "clear text", "cleartext", "no encryption"]},
    {"key": "tls_transit",
     "trigger": ["tls", "in transit", "transit"],
     "gr_link": ["tls", "transit"],
     "comply": ["tls 1.2", "tls 1.3", "encrypted in transit", "in transit (tls"],
     "violate": ["no tls", "plaintext transport", "http only", "unencrypted transit"]},
    {"key": "access_review",
     "trigger": ["recertif", "access review", "review access", "quarterly review"],
     "gr_link": ["recertif", "access review", "review"],
     "comply": ["quarterly access review", "access reviews scheduled",
                "reviewed quarterly", "quarterly access reviews", "access review scheduled"],
     "violate": ["no access review", "access is not reviewed", "not reviewed",
                 "never reviewed", "no review"]},
    {"key": "shared_admin",
     "trigger": ["shared", "generic account", "shared admin"],
     "gr_link": ["shared", "generic account"],
     "comply": ["named admin", "individual accounts", "no shared", "individually attributable"],
     "violate": ["shared admin", "single shared admin", "shared login",
                 "shared admin account", "generic account", "shared credentials"]},
    {"key": "admin_internet",
     "trigger": ["admin", "management console", "management interface"],
     "gr_link": ["admin interface", "administrative", "management"],
     "comply": ["behind vpn", "bastion", "not exposed", "restricted source"],
     "violate": ["exposed to the public internet", "published directly to the public internet",
                 "admin console exposed", "exposed to internet", "public internet"]},
    {"key": "firewall_deny",
     "trigger": ["firewall", "default-deny", "ports"],
     "gr_link": ["firewall", "default-deny", "default deny"],
     "comply": ["default-deny", "only required ports", "default deny"],
     "violate": ["allow any source", "allow-all", "any source ip", "allows any",
                 "allow all"]},
    {"key": "log_retention_pii",
     "trigger": ["log", "logs"],
     "gr_link": ["log", "retention", "purge", "retain"],
     "comply": ["purged within 90", "anonymised within 90", "anonymized within 90",
                "12-month retention", "retention configured", "logs are purged",
                "purge or anonymise"],
     "violate": ["retained indefinitely", "kept forever", "kept indefinitely",
                 "no purge", "retained for trend", "logs are retained indefinitely",
                 "no purge or anonymisation"]},
    {"key": "data_residency",
     "trigger": ["residency", "region", "eu/eea", "eu region"],
     "gr_link": ["residency", "eu", "region"],
     "comply": ["eu region", "eu/eea", "approved eu", "eu tenant", "within eu",
                "approved eu region", "hosted in an approved eu"],
     "violate": ["us region", "outside eu", "non-eu region", "stored in a us region"]},
    {"key": "backup_restore",
     "trigger": ["backup", "restore", "back up"],
     "gr_link": ["backup", "restore"],
     "comply": ["restore tested quarterly", "restores tested", "quarterly restore",
                "restore tests"],
     "violate": ["never restore-tested", "restores have never been tested",
                 "untested backups", "restores have never"]},
    {"key": "rto_rpo",
     "trigger": ["rto", "rpo", "recovery objective"],
     "gr_link": ["rto", "rpo"],
     "comply": ["rto and rpo", "rto/rpo agreed", "defined rto", "rto and rpo defined"],
     "violate": ["no rto", "no rpo", "no rto/rpo", "no rto or rpo",
                 "no rto/rpo has been agreed"]},
    {"key": "sast_dast",
     "trigger": ["sast", "dast", "security testing"],
     "gr_link": ["sast", "dast"],
     "comply": ["sast and dast run", "sast and dast", "security testing in the pipeline"],
     "violate": ["no security testing", "no sast", "no dast"]},
    {"key": "secrets",
     "trigger": ["secret", "hardcoded", "credential", "vault"],
     "gr_link": ["secret", "hardcoded", "credential"],
     "comply": ["managed vault", "secret store", "secrets pulled from a", "from a vault",
                "managed secret store"],
     "violate": ["hardcoded", "plaintext variables", "secrets in code"]},
    {"key": "code_review",
     "trigger": ["code review", "peer review", "pull request"],
     "gr_link": ["code review", "peer review"],
     "comply": ["peer review enforced", "peer review", "code review", "review before merge"],
     "violate": ["no code review", "no peer review", "unreviewed code"]},
    {"key": "least_privilege",
     "trigger": ["least privilege", "least-privilege"],
     "gr_link": ["least privilege", "least-privilege"],
     "comply": ["least-privilege", "least privilege", "least-privilege roles"],
     "violate": ["broad admin", "standing admin rights", "full admin to all"]},
    {"key": "password_length",
     "trigger": ["password"],
     "gr_link": ["password"],
     "comply": ["12 characters", "12-character", "at least 12", "strong password"],
     "violate": ["8-character password", "8 character password", "weak password",
                 "default password", "short password"]},
    {"key": "classification",
     "trigger": ["classif", "classified"],
     "gr_link": ["classif", "classified"],
     "comply": ["classified as", "internal-classified", "public-classified",
                "confidential", "restricted", "internal data", "data classified"],
     "violate": ["unclassified", "not classified", "no classification"]},
    {"key": "risk_assessment",
     "trigger": ["risk assessment", "power platform"],
     "gr_link": ["risk assessment"],
     "comply": ["risk assessment was completed", "risk assessment completed",
                "assessment was completed", "evidence linked"],
     "violate": ["no risk assessment", "without a risk assessment"]},
    {"key": "vuln_patch",
     "trigger": ["vuln", "patch", "remediat"],
     "gr_link": ["vuln", "patch", "remediat"],
     "comply": ["patched within 14 days", "critical vulns patched", "patched within",
                "remediated within 14"],
     "violate": ["unpatched", "not patched", "patched late"]},
    {"key": "supplier_assessment",
     "trigger": ["supplier", "third party", "third-party"],
     "gr_link": ["supplier", "third party", "third-party"],
     "comply": ["assessment completed", "supplier assessment completed",
                "security assessment completed"],
     "violate": ["no supplier assessment", "unassessed"]},
]

# Explicit "missing information" markers -> force a Gap when topic is relevant.
GAP_MARKERS = ["does not state", "does not specify", "not specified", "not stated",
               "does not mention", "unknown", "unstated", "missing", "unclear",
               "no information", "not clear whether", "does not say"]


def _contains_any(text, phrases):
    return [p for p in phrases if p in text]


def link_detectors(gdf):
    """Attach the set of detector keys relevant to each guardrail (by meaning)."""
    links = []
    for _, g in gdf.iterrows():
        blob = g["_text"].lower()
        keys = [d["key"] for d in DETECTORS if any(k in blob for k in d["gr_link"])]
        links.append(keys)
    gdf = gdf.copy()
    gdf["detector_keys"] = links
    return gdf


def evaluate_design(design_text, gdf):
    """
    Evaluate a design against the (deduped) guardrail model.
    Returns list of finding dicts. Every finding carries guardrail + clause cites.
    """
    text = " " + re.sub(r"\s+", " ", design_text.lower()) + " "
    det_by_key = {d["key"]: d for d in DETECTORS}
    findings = []

    for _, g in gdf.iterrows():
        keys = g.get("detector_keys", [])
        if not keys:
            continue
        for k in keys:
            det = det_by_key[k]
            # Is the topic in scope for this design?
            trig_hits = _contains_any(text, det["trigger"])
            if not trig_hits:
                continue
            viol_hits = _contains_any(text, det["violate"])
            comp_hits = _contains_any(text, det["comply"])
            gap_hits = _contains_any(text, GAP_MARKERS)

            if viol_hits:
                verdict, evidence = "Non-compliant", viol_hits
            elif gap_hits and not comp_hits:
                verdict, evidence = "Gap", gap_hits
            elif comp_hits:
                verdict, evidence = "Compliant", comp_hits
            else:
                # Topic mentioned but neither satisfied, violated, nor explicitly
                # flagged as missing -> informational only, not a fabricated gap.
                verdict, evidence = "Not applicable", ["topic mentioned; no explicit evidence"]

            if verdict == "Not applicable":
                continue

            findings.append({
                "guardrail_id": g["guardrail_id"],
                "statement": g["statement"],
                "domain": g["domain"],
                "category": g["category"],
                "severity": g["severity"],
                "detector": k,
                "verdict": verdict,
                "evidence": "; ".join(evidence),
                "trigger_matched": "; ".join(trig_hits),
                "source_clause_ids": ", ".join(g["source_clause_ids"]),
                "source_document_ids": ", ".join(g["source_document_ids"]),
                "clause_ref": g.get("clause_ref", ""),
                "requirement": g["requirement"],
            })

    # De-duplicate: keep worst verdict per guardrail (Non-compliant > Gap > Compliant)
    worst = {"Non-compliant": 3, "Gap": 2, "Compliant": 1}
    best_per_gr = {}
    for f in findings:
        gid = f["guardrail_id"]
        if gid not in best_per_gr or worst[f["verdict"]] > worst[best_per_gr[gid]["verdict"]]:
            best_per_gr[gid] = f
    result = sorted(best_per_gr.values(),
                    key=lambda x: (-worst[x["verdict"]], -SEVERITY_ORDER[x["severity"]]))
    return result


# ============================================================================= #
# 8. REMEDIATION PLANNER
# ============================================================================= #
REMEDIATION_HINTS = {
    "public_storage": "Enable public-access-block; grant bucket access explicitly via least-privilege policy.",
    "mfa": "Enforce MFA on all privileged / remote / internet-facing logins; log MFA events.",
    "encryption_at_rest": "Enable storage encryption (AES-256) for Confidential/Restricted and all PII stores.",
    "tls_transit": "Enforce TLS 1.2+ on all endpoints; disable legacy ciphers and plaintext.",
    "access_review": "Schedule quarterly access recertification by resource owners; retain evidence.",
    "shared_admin": "Replace shared admin logins with individually attributable named accounts.",
    "admin_internet": "Move admin/management interfaces behind VPN/bastion; restrict source IPs.",
    "firewall_deny": "Apply default-deny firewall posture; open only required ports/protocols.",
    "log_retention_pii": "Set a 90-day purge/anonymisation job for logs with personal data (unless legal hold).",
    "data_residency": "Host EU personal data in approved EU/EEA regions or apply a valid transfer mechanism.",
    "backup_restore": "Run quarterly restore tests; record results to prove recoverability.",
    "rto_rpo": "Agree and document RTO/RPO with the business owner for the critical system.",
    "sast_dast": "Integrate SAST and DAST gates in CI/CD; fail builds on critical findings.",
    "secrets": "Retrieve secrets from a managed vault at runtime; scan repos for hardcoded secrets.",
    "code_review": "Enforce peer pull-request review before merge to protected branches.",
    "least_privilege": "Scope permissions to the task; remove standing/broad admin rights.",
    "password_length": "Set privileged-account minimum length to 12 chars; enforce complexity/rotation.",
    "classification": "Classify data (Public/Internal/Confidential/Restricted) before storage or processing.",
    "risk_assessment": "Complete the Power Platform risk assessment and link the evidence.",
    "vuln_patch": "Remediate critical vulnerabilities within 14 days; track the SLA.",
    "supplier_assessment": "Complete the supplier security assessment before onboarding; add contractual controls.",
}


def build_remediation(findings):
    rows = []
    priority = {"High": 1, "Medium": 2, "Low": 3}
    for f in findings:
        if f["verdict"] == "Compliant":
            continue
        rows.append({
            "Priority": priority[f["severity"]],
            "Severity": f["severity"],
            "Verdict": f["verdict"],
            "Guardrail": f["guardrail_id"],
            "Requirement": f["statement"],
            "Recommended action": REMEDIATION_HINTS.get(
                f["detector"], "Review against the cited clause and remediate."),
            "Cited clauses": f["source_clause_ids"],
            "Cited documents": f["source_document_ids"],
        })
    rdf = pd.DataFrame(rows)
    if not rdf.empty:
        rdf = rdf.sort_values(["Priority", "Verdict"]).reset_index(drop=True)
    return rdf


# ============================================================================= #
# 9. MULTI-SHEET AUDIT EXPORT
# ============================================================================= #
def build_audit_export(gdf_raw, ddf, conflicts, findings, remediation, design_name):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        # Sheet 1: summary
        summary = pd.DataFrame({
            "Metric": ["Generated", "Design evaluated", "Raw guardrails",
                       "Deduplicated guardrails", "Merged clusters",
                       "Numeric conflicts", "Findings", "Non-compliant", "Gaps",
                       "Compliant"],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                design_name or "-",
                len(gdf_raw), len(ddf),
                int(ddf["is_merged"].sum()) if "is_merged" in ddf else 0,
                len(conflicts), len(findings),
                sum(1 for f in findings if f["verdict"] == "Non-compliant"),
                sum(1 for f in findings if f["verdict"] == "Gap"),
                sum(1 for f in findings if f["verdict"] == "Compliant"),
            ],
        })
        summary.to_excel(xw, sheet_name="Summary", index=False)

        # Sheet 2: guardrail model
        gm = ddf.copy()
        for col in ["source_clause_ids", "source_document_ids", "merged_from", "orphan_refs"]:
            if col in gm:
                gm[col] = gm[col].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
        gm.drop(columns=["_text", "detector_keys"], errors="ignore").to_excel(
            xw, sheet_name="Guardrail_Model", index=False)

        # Sheet 3: conflicts
        crows = []
        for c in conflicts:
            for p in c["positions"]:
                crows.append({
                    "Measure": c["measure"], "Value": p["value_raw"],
                    "Normalised": p["value_norm"], "Guardrail": p["guardrail_id"],
                    "Clauses": ", ".join(p["clauses"]),
                    "Documents": ", ".join(p["documents"]),
                    "Resolution": c["resolution"],
                })
        pd.DataFrame(crows or [{"Measure": "None detected"}]).to_excel(
            xw, sheet_name="Conflicts", index=False)

        # Sheet 4: findings
        pd.DataFrame(findings or [{"note": "no findings"}]).to_excel(
            xw, sheet_name="Findings", index=False)

        # Sheet 5: remediation
        (remediation if not remediation.empty else
         pd.DataFrame([{"note": "nothing to remediate"}])).to_excel(
            xw, sheet_name="Remediation_Plan", index=False)
    buf.seek(0)
    return buf.getvalue()


# ============================================================================= #
# 10. STREAMLIT UI
# ============================================================================= #
def sev_badge(sev):
    return f"<span style='background:{SEVERITY_COLOR[sev]};color:white;padding:2px 8px;border-radius:10px;font-size:0.8em'>{sev}</span>"


def verdict_badge(v):
    return f"<span style='background:{VERDICT_COLOR[v]};color:white;padding:2px 8px;border-radius:10px;font-size:0.85em'>{v}</span>"


def main():
    st.title("🛡️ GuardrailIQ — Deterministic GRC Compliance Reviewer")
    st.caption("Governance Excel → traceable guardrail model → design evaluation. "
               "No LLM · fully explainable · every finding cites its clause.")

    ss = st.session_state
    ss.setdefault("built", False)

    # ---------------- Sidebar: ingest & build ----------------
    with st.sidebar:
        st.header("1 · Load governance workbook")
        up = st.file_uploader("GRC Excel (.xlsx)", type=["xlsx"])
        sim = st.slider("Dedup similarity threshold", 0.15, 0.60, 0.26, 0.01,
                        help="Higher = stricter (fewer merges).")
        if up and st.button("⚙️ Build guardrail model", type="primary"):
            docs, clauses, designs, report = load_workbook(up.read())
            gdf_raw = build_guardrails(clauses, docs)
            ddf, clusters, simm, ids = deduplicate(gdf_raw, sim)
            ddf = link_detectors(ddf)
            gdf_raw = link_detectors(gdf_raw)
            conflicts = detect_conflicts(ddf)
            ss.update(dict(built=True, docs=docs, clauses=clauses, designs=designs,
                           report=report, gdf_raw=gdf_raw, ddf=ddf, clusters=clusters,
                           conflicts=conflicts, sim=sim, findings=[], design_name=""))
            st.success(f"Built {len(ddf)} guardrails from {len(gdf_raw)} clauses.")

        st.divider()
        st.markdown("**Load sample data** if you don't have a file handy.")
        if st.button("📥 Use bundled sample workbook"):
            try:
                with open("/mnt/user-data/uploads/GuardrailIQ_Dataset_v2.xlsx", "rb") as fh:
                    data = fh.read()
                docs, clauses, designs, report = load_workbook(data)
                gdf_raw = build_guardrails(clauses, docs)
                ddf, clusters, simm, ids = deduplicate(gdf_raw, sim)
                ddf = link_detectors(ddf); gdf_raw = link_detectors(gdf_raw)
                conflicts = detect_conflicts(ddf)
                ss.update(dict(built=True, docs=docs, clauses=clauses, designs=designs,
                               report=report, gdf_raw=gdf_raw, ddf=ddf, clusters=clusters,
                               conflicts=conflicts, sim=sim, findings=[], design_name=""))
                st.success(f"Sample loaded · {len(ddf)} guardrails.")
            except Exception as e:
                st.error(f"Sample not found: {e}")

    if not ss.built:
        st.info("⬅️ Upload the governance workbook (or load the sample) and click "
                "**Build guardrail model** to begin.")
        _render_architecture_tab()
        return

    tabs = st.tabs(["📊 Dashboard", "🧱 Guardrail model", "🔀 Conflicts",
                    "🔎 Evaluate design", "🛠️ Remediation", "📤 Audit export",
                    "🏗️ Architecture & demo"])

    with tabs[0]:
        _render_dashboard(ss)
    with tabs[1]:
        _render_model(ss)
    with tabs[2]:
        _render_conflicts(ss)
    with tabs[3]:
        _render_evaluate(ss)
    with tabs[4]:
        _render_remediation(ss)
    with tabs[5]:
        _render_export(ss)
    with tabs[6]:
        _render_architecture_tab()


def _render_dashboard(ss):
    ddf, gdf_raw = ss.ddf, ss.gdf_raw
    c = st.columns(5)
    c[0].metric("Source clauses", len(ss.clauses))
    c[1].metric("Raw guardrails", len(gdf_raw))
    c[2].metric("Deduplicated", len(ddf))
    c[3].metric("Merged clusters", int(ddf["is_merged"].sum()))
    c[4].metric("Numeric conflicts", len(ss.conflicts))

    col1, col2 = st.columns(2)
    with col1:
        sev = ddf["severity"].value_counts().reindex(["High", "Medium", "Low"]).fillna(0)
        fig = px.bar(x=sev.index, y=sev.values, color=sev.index,
                     color_discrete_map=SEVERITY_COLOR,
                     labels={"x": "Severity", "y": "Guardrails"},
                     title="Guardrails by severity")
        fig.update_layout(showlegend=False, height=340)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        dom = ddf["domain"].value_counts()
        fig = px.pie(values=dom.values, names=dom.index, title="Guardrails by domain", hole=0.4)
        fig.update_layout(height=340)
        st.plotly_chart(fig, use_container_width=True)

    orphans = ddf[ddf["orphan_refs"].apply(lambda x: len(x) > 0)]
    if not orphans.empty:
        st.warning(f"⚠️ {len(orphans)} guardrail(s) contain **orphan (broken) references** "
                   "to documents that do not exist in the workbook — flagged as untraceable.")
        st.dataframe(orphans[["guardrail_id", "statement", "orphan_refs"]],
                     use_container_width=True, hide_index=True)


def _render_model(ss):
    ddf = ss.ddf
    st.subheader("Deduplicated, traceable guardrail model")
    st.caption("Each row is one canonical guardrail. Merged rows preserve ALL source clauses.")
    only_merged = st.checkbox("Show only merged (consolidated) guardrails")
    view = ddf[ddf["is_merged"]] if only_merged else ddf
    show = view[["guardrail_id", "statement", "domain", "category", "severity",
                 "source_clause_ids", "source_document_ids", "is_merged"]].copy()
    show["source_clause_ids"] = show["source_clause_ids"].apply(", ".join)
    show["source_document_ids"] = show["source_document_ids"].apply(", ".join)
    st.dataframe(show, use_container_width=True, hide_index=True, height=430)

    if ss.clusters:
        st.subheader("🔗 De-duplication clusters (TF-IDF + cosine + NetworkX)")
        st.caption("Overlapping requirements consolidated into one guardrail; "
                   "every source clause is retained for traceability.")
        for cl in ss.clusters:
            with st.expander(f"{cl['canonical']} · avg similarity {cl['avg_similarity']} · "
                             f"{len(cl['members'])} merged → {cl['statement']}"):
                st.write(f"**Merged guardrails:** {', '.join(cl['members'])}")
                st.write(f"**Preserved source clauses:** {', '.join(cl['clauses'])}")


def _render_conflicts(ss):
    st.subheader("🔀 Cross-document numeric-threshold conflicts")
    st.caption("Different numeric values for the same requirement are SURFACED, "
               "never auto-resolved. A human must decide.")
    if not ss.conflicts:
        st.success("No numeric-threshold conflicts detected.")
        return
    for c in ss.conflicts:
        st.markdown(f"### ⚠️ {c['measure']}")
        st.markdown(f"**Distinct values found:** {c['distinct_values']} · "
                    f"**Resolution:** :red[{c['resolution']}]")
        rows = [{"Guardrail": p["guardrail_id"], "Stated value": p["value_raw"],
                 "Normalised": p["value_norm"], "Domain": p["domain"],
                 "Cited clauses": ", ".join(p["clauses"]),
                 "Cited documents": ", ".join(p["documents"])}
                for p in c["positions"]]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        st.divider()


def _render_evaluate(ss):
    st.subheader("🔎 Evaluate a design against the guardrail model")
    src = st.radio("Design source", ["Upload file (DOCX/PDF/PPTX/TXT)",
                                      "Pick a sample design from the workbook",
                                      "Paste text"], horizontal=True)
    design_text, design_name = "", ""

    if src.startswith("Upload"):
        f = st.file_uploader("Design document", type=["docx", "pdf", "pptx", "txt"])
        if f:
            design_text = extract_text(f)
            design_name = f.name
    elif src.startswith("Pick"):
        designs = ss.designs
        if designs.empty:
            st.info("No Design_Descriptions sheet found in the workbook.")
        else:
            opt = st.selectbox("Sample design",
                               designs["Design_ID"] + " — " + designs["Design_Title"])
            did = opt.split(" — ")[0]
            row = designs[designs["Design_ID"] == did].iloc[0]
            design_text = f"{row['Design_Title']}. {row['Design_Description']}"
            design_name = did
            with st.expander("Design text"):
                st.write(design_text)
            if row.get("Expected_Outcome"):
                st.caption(f"Dataset self-check label (not used by the engine): "
                           f"**{row['Expected_Outcome']}**")
    else:
        design_text = st.text_area("Paste design description", height=180)
        design_name = "pasted-design"

    if design_text and st.button("▶️ Run evaluation", type="primary"):
        findings = evaluate_design(design_text, ss.ddf)
        ss.findings = findings
        ss.design_name = design_name

    if ss.get("findings"):
        findings = ss.findings
        nc = sum(1 for f in findings if f["verdict"] == "Non-compliant")
        gp = sum(1 for f in findings if f["verdict"] == "Gap")
        cp = sum(1 for f in findings if f["verdict"] == "Compliant")
        overall = ("Non-compliant" if nc else "Gap" if gp else "Compliant")
        st.markdown(f"## Overall: {verdict_badge(overall)}", unsafe_allow_html=True)
        m = st.columns(3)
        m[0].metric("❌ Non-compliant", nc)
        m[1].metric("⚠️ Gaps", gp)
        m[2].metric("✅ Compliant", cp)

        for f in findings:
            with st.container(border=True):
                top = st.columns([3, 1])
                top[0].markdown(
                    f"**{f['guardrail_id']} — {f['statement']}**  "
                    f"{sev_badge(f['severity'])}", unsafe_allow_html=True)
                top[1].markdown(verdict_badge(f["verdict"]), unsafe_allow_html=True)
                st.markdown(
                    f"*Evidence in design:* `{f['evidence']}`  \n"
                    f"*Detector:* `{f['detector']}` · *Domain:* {f['domain']}  \n"
                    f"📌 **Cited guardrail:** {f['guardrail_id']} · "
                    f"**Source clause(s):** {f['source_clause_ids']} · "
                    f"**Document(s):** {f['source_document_ids']}")
                with st.expander("Requirement text (traceability)"):
                    st.write(f["requirement"])


def _render_remediation(ss):
    st.subheader("🛠️ Remediation planner")
    if not ss.get("findings"):
        st.info("Run a design evaluation first (Evaluate design tab).")
        return
    rdf = build_remediation(ss.findings)
    if rdf.empty:
        st.success("No remediation required — all evaluated guardrails are compliant. 🎉")
        return
    st.caption("Prioritised by severity. Each action cites the guardrail clause behind it.")
    show = rdf.copy()
    show["Cited clauses"] = show["Cited clauses"].apply(
        lambda x: ", ".join(x) if isinstance(x, list) else x)
    show["Cited documents"] = show["Cited documents"].apply(
        lambda x: ", ".join(x) if isinstance(x, list) else x)
    st.dataframe(show, use_container_width=True, hide_index=True)


def _render_export(ss):
    st.subheader("📤 Multi-sheet audit export")
    st.caption("Summary · Guardrail_Model · Conflicts · Findings · Remediation_Plan")
    remediation = build_remediation(ss.get("findings", []))
    data = build_audit_export(ss.gdf_raw, ss.ddf, ss.conflicts,
                              ss.get("findings", []), remediation,
                              ss.get("design_name", ""))
    st.download_button("⬇️ Download audit workbook (.xlsx)", data=data,
                       file_name=f"GuardrailIQ_Audit_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    st.dataframe(ss.ddf[["guardrail_id", "statement", "severity"]].head(20),
                 use_container_width=True, hide_index=True)


def _render_architecture_tab():
    st.subheader("🏗️ Architecture & how it works")
    st.markdown("""
**Deterministic, explainable pipeline — no external LLM.**

```
┌──────────────┐   flexible    ┌───────────────┐   severity     ┌──────────────────┐
│ Governance   │──alias map──▶ │ Normalised    │──tag+category▶ │ Raw guardrails   │
│ Excel (.xlsx)│               │ tables        │                │ (1 per clause)   │
└──────────────┘               └───────────────┘                └────────┬─────────┘
                                                                          │ TF-IDF + cosine
                                                                          ▼ + NetworkX
┌──────────────┐  keyword-      ┌───────────────┐   numeric     ┌──────────────────┐
│ Design doc   │──evidence────▶ │ Findings:     │◀──threshold───│ Deduplicated     │
│ DOCX/PDF/... │   matching     │ Compliant /   │   conflicts   │ guardrail model  │
└──────────────┘               │ Non-comp/Gap  │  (surfaced)    │ (sources kept)   │
                               │ + citations   │                └──────────────────┘
                               └───────┬───────┘
                                       ▼
                 Dashboard · Remediation planner · Multi-sheet audit export
```

**Why it is trustworthy**
- **Traceability gate:** every finding carries its guardrail ID *and* source clause/document IDs.
- **Meaning, not IDs:** severity, dedup, conflicts and evaluation all work by text/number
  analysis, so a *modified copy* of the dataset (new IDs, new clauses) still works.
- **No auto-resolve:** cross-document numeric conflicts are surfaced for a human, never silently picked.
- **Gaps stay gaps:** when a design omits required info, the engine returns *Gap*, not a fabricated pass/fail.
""")


if __name__ == "__main__":
    main()
