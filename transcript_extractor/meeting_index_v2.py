import os
import re
import io
import sys
import time
import logging
import argparse
import urllib.parse
from collections import defaultdict
from datetime import datetime
from pathlib import Path
 
try:
    from dotenv import load_dotenv
    _env = Path(__file__).resolve().parent / ".env"
    if not _env.exists():
        _env = Path(__file__).resolve().parent.parent / ".env"
    load_dotenv(dotenv_path=_env, override=True)
except ImportError:
    pass
 
import requests
from msal import ConfidentialClientApplication
from docx import Document   # python-docx (already a dependency of the extractor)
 
logger = logging.getLogger(__name__)
 
 
TENANT_ID = os.environ.get("TENANT_ID", "")
CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
DRIVE_ID = os.environ.get("SHAREPOINT_DRIVE_ID", "")
# Default to the YEAR root so month rollover (July+) does not silently stop scanning.
TRANSCRIPTS_ROOT = os.environ.get("TRANSCRIPTS_ROOT", "Transcripts/2026")
DATAVERSE_URL = os.environ.get("DATAVERSE_URL", "").rstrip("/")
DATAVERSE_TABLE = "crb6b_table1s"
 
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
 
# Longest / most specific acronyms first so HUSAM wins over SAM, etc.
# Values are STORED in Dataverse and used by the Copilot topics for filtering.
MEETING_TYPES = [
    (["HUSAM", "HU STRATEGIC"], "HUSAM"),
    (["DOF", "DAILY OPERATIONAL"], "DOF"),
    (["OCC"], "OCC"),
    (["WEEKEND"], "Weekend Call"),
    (["EOD"], "EOD"),
    (["LEADERSHIP TEAM"], "Leadership Team"),
    (["SAM", "STRATEGIC ALIGNMENT"], "SAM"),
]
 
DATE_RE_DASH = re.compile(r"(20\d{2})-(\d{2})-(\d{2})")
DATE_RE_COMPACT = re.compile(r"(20\d{2})(\d{2})(\d{2})")
 
# Must stay <= the crb6b_summary column's max length. Single-line text maxes at
# 4000; set DIGEST_MAX_CHARS=8000 (etc.) if you make the column Multiple lines of text.
DIGEST_MAX_CHARS = int(os.environ.get("DIGEST_MAX_CHARS", "3900"))
 
# Reasoning models (gpt-5*, o*) spend hidden "reasoning tokens" out of the SAME
# max_completion_tokens budget. At 2000, long transcripts exhausted the whole
# budget on reasoning and the API returned EMPTY content with finish_reason=
# "length" and NO error -> blank summaries silently inserted into Dataverse.
# Only tokens actually used are billed, so a big ceiling costs nothing extra.
DIGEST_MAX_COMPLETION_TOKENS = int(os.environ.get("DIGEST_MAX_COMPLETION_TOKENS", "8000"))
# Optional: AZURE_OPENAI_REASONING_EFFORT=low cuts reasoning burn on gpt-5
# models (requires an api-version that supports it, e.g. 2025-01-01-preview+).
DIGEST_REASONING_EFFORT = os.environ.get("AZURE_OPENAI_REASONING_EFFORT", "").strip()
 
DIGEST_PROMPT = (
    "You are extracting facts from a South Texas Project meeting transcript. It may "
    "be any meeting type — DOF, HUSAM, SAM, OCC, Weekend Call, PORC, a staff or "
    "leadership meeting, etc. Output a SHORT digest with ONLY what is explicitly "
    "stated. Use these sections and OMIT any that do not apply to this meeting:\n"
    "KEY TOPICS: <main items discussed>\n"
    "OPERATIONAL CONCERNS: <equipment/status issues, if any>\n"
    "DECISIONS: <bullets>\n"
    "ACTION ITEMS: <bullets, include owner if stated>\n"
    "RISK/STATUS: <unit MWe, risk colors, CR numbers, RCS leakage, if stated>\n"
    "Do not invent CR numbers, names, dates, or values. Keep under 250 words."
)
 
 
def _get_graph_token():
    if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID]):
        raise RuntimeError("Missing CLIENT_ID, CLIENT_SECRET, or TENANT_ID in .env")
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/" + TENANT_ID,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError("Graph auth failed: " + str(result.get("error_description")))
    return result["access_token"]
 
 
def _get_dataverse_token():
    if not DATAVERSE_URL:
        raise RuntimeError("DATAVERSE_URL not set in .env")
    if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID]):
        raise RuntimeError("Missing CLIENT_ID, CLIENT_SECRET, or TENANT_ID in .env")
    dataverse_resource = DATAVERSE_URL.split("/api/")[0]
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/" + TENANT_ID,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=[dataverse_resource + "/.default"])
    if "access_token" not in result:
        raise RuntimeError("Dataverse auth failed: " + str(result.get("error_description")))
    return result["access_token"]
 
 
def _graph_get(url, token, max_retries=5):
    headers = {"Authorization": "Bearer " + token, "Accept": "application/json"}
    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=60)
        except requests.exceptions.RequestException as e:
            last_err = e
            logger.warning("Graph GET network error (attempt %d/%d): %s", attempt, max_retries, e)
            time.sleep(min(2 ** attempt, 30))
            continue
 
        if resp.status_code == 429 or resp.status_code in (500, 502, 503, 504):
            retry_after = int(resp.headers.get("Retry-After", min(2 ** attempt, 30)))
            logger.warning("Graph GET %d (attempt %d/%d). Waiting %ds...",
                           resp.status_code, attempt, max_retries, retry_after)
            last_err = RuntimeError("HTTP %d" % resp.status_code)
            time.sleep(retry_after)
            continue
 
        resp.raise_for_status()
        return resp.json()
 
    raise RuntimeError(
        "SharePoint listing failed for %s after %d attempts (%s). "
        "Aborting so a partial sync is not mistaken for complete."
        % (url, max_retries, last_err)
    )
 
 
def _graph_get_bytes(url, token, max_retries=5):
    """Download raw file bytes (the .docx) with the same retry policy."""
    headers = {"Authorization": "Bearer " + token}
    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=120)
        except requests.exceptions.RequestException as e:
            last_err = e
            time.sleep(min(2 ** attempt, 30))
            continue
        if resp.status_code in (429, 500, 502, 503, 504):
            retry_after = int(resp.headers.get("Retry-After", min(2 ** attempt, 30)))
            last_err = RuntimeError("HTTP %d" % resp.status_code)
            time.sleep(retry_after)
            continue
        resp.raise_for_status()
        return resp.content
    raise RuntimeError("File download failed for %s (%s)" % (url, last_err))
 
 
def _parse_graph_dt(value):
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:19], "%Y-%m-%dT%H:%M:%S")
    except (ValueError, TypeError):
        return None
 
 
def _list_docx_files(token, root):
    files = []
    url = GRAPH_BASE + "/drives/" + DRIVE_ID + "/root:/" + root + ":/children"
    _recurse(token, url, files, root)
    return files
 
 
def _recurse(token, url, files, root):
    data = _graph_get(url, token)  # raises on hard failure -- intentional
 
    for item in data.get("value", []):
        if item.get("folder"):
            child_url = GRAPH_BASE + "/drives/" + DRIVE_ID + "/items/" + item["id"] + "/children"
            _recurse(token, child_url, files, root)
            continue
 
        nm = item.get("name", "")
        if item.get("file") and nm.lower().endswith(".docx") and not nm.startswith("~$"):
            web_url = item.get("webUrl", "")
            if not web_url:
                web_url = ("https://stpgweb.sharepoint.com/sites/ChatSTP/"
                           "Shared%20Documents/" + root + "/" + nm)
            files.append({
                "name": nm,
                "id": item["id"],
                "webUrl": web_url,
                "created": _parse_graph_dt(item.get("createdDateTime")),
                "modified": _parse_graph_dt(item.get("lastModifiedDateTime")),
            })
 
    next_link = data.get("@odata.nextLink")
    if next_link:
        _recurse(token, next_link, files, root)
 
 
def _extract_docx_text(content_bytes, stop_marker="Raw Transcript (VTT)", max_chars=30000):
    """Clean transcript text from a .docx, dropping the duplicated raw VTT dump."""
    try:
        doc = Document(io.BytesIO(content_bytes))
    except Exception as e:
        logger.warning("docx parse failed: %s", e)
        return ""
    lines = []
    for p in doc.paragraphs:
        t = (p.text or "").rstrip()
        if stop_marker.lower() in t.lower():
            break
        lines.append(t)
    return "\n".join(lines).strip()[:max_chars]
 
 
def _fetch_transcript_text(graph_token, item_id, max_chars=30000):
    if not item_id:
        return ""
    url = GRAPH_BASE + "/drives/" + DRIVE_ID + "/items/" + item_id + "/content"
    try:
        return _extract_docx_text(_graph_get_bytes(url, graph_token), max_chars=max_chars)
    except Exception as e:
        logger.warning("Could not download content for item %s: %s", item_id, e)
        return ""
 
 
def _make_digest(transcript_text):
    if not transcript_text:
        return ""
    text = transcript_text[:30000]
 
    # Azure OpenAI
    if os.environ.get("AZURE_OPENAI_ENDPOINT") and os.environ.get("AZURE_OPENAI_KEY"):
        try:
            from openai import AzureOpenAI
            client = AzureOpenAI(
                api_key=os.environ["AZURE_OPENAI_KEY"],
                api_version=os.environ.get("AZURE_OPENAI_API_VERSION", "2024-10-21"),
                azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
            )
            kwargs = dict(
                model=os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-5-mini"),
                messages=[{"role": "system", "content": DIGEST_PROMPT},
                          {"role": "user", "content": text}],
                max_completion_tokens=DIGEST_MAX_COMPLETION_TOKENS,
            )
            if DIGEST_REASONING_EFFORT:
                kwargs["reasoning_effort"] = DIGEST_REASONING_EFFORT
            resp = client.chat.completions.create(**kwargs)
            choices = getattr(resp, "choices", None) or []
            if not choices:
                logger.warning("Azure OpenAI returned no choices for this row")
                return ""
            finish = getattr(choices[0], "finish_reason", None)
            content = getattr(choices[0].message, "content", None)
            digest = (content or "").strip()
            if not digest:
                logger.warning(
                    "Azure OpenAI returned EMPTY content (finish_reason=%s). "
                    "Reasoning tokens likely consumed the completion budget -- "
                    "raise DIGEST_MAX_COMPLETION_TOKENS -- or the content filter fired.",
                    finish)
                return ""
            return digest[:DIGEST_MAX_CHARS]
        except Exception as e:
            logger.warning("Azure OpenAI digest failed: %s", e)
            return ""
 
    # Claude (Anthropic)
    if os.environ.get("ANTHROPIC_API_KEY"):
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
            msg = client.messages.create(
                model=os.environ.get("ANTHROPIC_MODEL", "claude-haiku-4-5-20251001"),
                max_tokens=600,
                system=DIGEST_PROMPT,
                messages=[{"role": "user", "content": text}],
            )
            digest = (msg.content[0].text or "").strip()
            if not digest:
                logger.warning("Claude returned empty content for this row")
                return ""
            return digest[:DIGEST_MAX_CHARS]
        except Exception as e:
            logger.warning("Claude digest failed: %s", e)
            return ""
 
    logger.warning("No LLM configured (set AZURE_OPENAI_* or ANTHROPIC_API_KEY) — "
                   "storing empty summary")
    return ""
 
 
def _extract_date(text):
    m = DATE_RE_DASH.search(text)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = DATE_RE_COMPACT.search(text)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    return None
 
 
def _detect_type(text):
    upper = re.sub(r"[_\-]+", " ", text).upper()
    for keys, val in MEETING_TYPES:
        if any(k in upper for k in keys):
            return val
    return "Other"
 
 
def _clean_name(name):
    clean = name
    clean = DATE_RE_DASH.sub(" ", clean)
    clean = DATE_RE_COMPACT.sub(" ", clean)
    clean = re.sub(r"\d{6}\s*UTC", " ", clean, flags=re.IGNORECASE)
    clean = re.sub(r"\bUTC\b", " ", clean, flags=re.IGNORECASE)
    clean = re.sub(r"[Tt]\d{6}", " ", clean)
    clean = re.sub(r"(?<!\d)(\d{6}|\d{4})(?!\d)", " ", clean)
    clean = re.sub(r"auto\s*record", " ", clean, flags=re.IGNORECASE)
    clean = re.sub(r"meeting[\s_-]*recording", " ", clean, flags=re.IGNORECASE)
    clean = re.sub(r"\brecording\b", " ", clean, flags=re.IGNORECASE)
    clean = clean.replace("_", " ").replace("-", " ")
    clean = re.sub(r"\s+", " ", clean).strip(" -_")
    return clean
 
 
def _parse_filename(filename, fallback_dt=None):
    name = re.sub(r"\.docx$", "", filename, flags=re.IGNORECASE).strip()
 
    dt = _extract_date(name)
    if dt is None:
        dt = fallback_dt
    if dt is None:
        return None
 
    meeting_type = _detect_type(name)
    clean = _clean_name(name)
 
    if meeting_type != "Other":
        meeting_name = meeting_type
        if meeting_type == "OCC":
            shift = re.search(r"(?<!\d)(0600|1800)(?!\d)", name)
            if shift:
                meeting_name = "OCC (%s)" % shift.group(1)
        unnamed = False
    elif clean:
        meeting_name = clean
        unnamed = False
    else:
        meeting_name = "Unknown Meeting"
        unnamed = True
 
    return {
        "meeting_name": meeting_name,
        "meeting_date": dt.strftime("%Y-%m-%d"),
        "month": dt.strftime("%B"),
        "year": dt.year,
        "meeting_type": meeting_type,
        "unnamed": unnamed,
    }
 
 
def _norm_filename(value):
    if not value:
        return ""
    # SharePoint "_layouts/15/Doc.aspx" redirect URLs carry the real filename in
    # the ?file= query param — read that first.
    try:
        qs = urllib.parse.parse_qs(urllib.parse.urlparse(value).query)
        if qs.get("file"):
            return urllib.parse.unquote(qs["file"][0]).strip().lower()
    except Exception:
        pass
    # Otherwise use the last path segment (clean path URLs / bare filenames).
    v = value.split("?")[0].rstrip("/").split("/")[-1]
    v = urllib.parse.unquote(v)
    return v.strip().lower()
 
 
def _dv_headers(dv_token, write=False):
    h = {
        "Authorization": "Bearer " + dv_token,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    if write:
        h["Content-Type"] = "application/json"
    return h
 
 
def _get_existing_rows(dv_token, select="crb6b_fileurl,crb6b_meetingname,crb6b_meetingdate,crb6b_meetingtype"):
    """Fetch ALL rows (handles paging). Returns list of dicts."""
    url = (DATAVERSE_URL + "/api/data/v9.2/" + DATAVERSE_TABLE + "?$select=" + select)
    headers = _dv_headers(dv_token)
    rows = []
    page = 0
    while url:
        page += 1
        try:
            resp = requests.get(url, headers=headers, timeout=30)
        except requests.exceptions.RequestException as e:
            logger.error("Dataverse fetch error on page %d: %s", page, e)
            raise
        if resp.status_code == 404:
            raise RuntimeError("Dataverse table not found -- check DATAVERSE_TABLE name")
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 5))
            logger.warning("Dataverse throttled. Waiting %ds...", retry_after)
            time.sleep(retry_after)
            continue
        resp.raise_for_status()
 
        data = resp.json()
        rows.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return rows
 
 
def _existing_filename_set(rows):
    s = set()
    for r in rows:
        k = _norm_filename(r.get("crb6b_fileurl", ""))
        if k:
            s.add(k)
    return s
 
 
def _row_exists_for_file(dv_token, file_url):
    """True if a row with this exact crb6b_fileurl already exists.
    Used to detect an insert that COMMITTED server-side even though the
    client saw a timeout -- blindly retrying the POST would duplicate it
    (this is how twin rows like 5/5 DOF x2 and 4/11 OCC x2 get created)."""
    if not file_url:
        return False
    try:
        flt = "crb6b_fileurl eq '" + file_url.replace("'", "''") + "'"
        resp = requests.get(
            DATAVERSE_URL + "/api/data/v9.2/" + DATAVERSE_TABLE,
            headers=_dv_headers(dv_token),
            params={"$select": "crb6b_table1id", "$top": "1", "$filter": flt},
            timeout=30)
        return resp.status_code == 200 and len(resp.json().get("value", [])) > 0
    except Exception:
        return False  # unknown -> fall back to normal retry behaviour
 
 
def _upsert_record(dv_token, metadata, file_url, digest="", transcript_text="", max_retries=5):
    url = DATAVERSE_URL + "/api/data/v9.2/" + DATAVERSE_TABLE
    headers = _dv_headers(dv_token, write=True)
    body = {
        "crb6b_meetingname": metadata["meeting_name"],
        "crb6b_meetingdate": metadata["meeting_date"],
        "crb6b_month": metadata["month"],
        "crb6b_year": metadata["year"],
        "crb6b_meetingtype": metadata["meeting_type"],
        "crb6b_fileurl": file_url,
        "crb6b_summary": digest,
        "crb6b_transcripttext": transcript_text,
    }
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.post(url, headers=headers, json=body, timeout=30)
        except requests.exceptions.RequestException as e:
            logger.warning("Attempt %d/%d: network error inserting: %s", attempt, max_retries, e)
            # The POST may have landed despite the client-side error/timeout.
            if _row_exists_for_file(dv_token, file_url):
                logger.info("Row already exists after network error -- treating as success")
                return True
            if attempt < max_retries:
                time.sleep(min(2 ** attempt, 30))
                continue
            return False
 
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 5))
            logger.warning("Attempt %d/%d: throttled. Waiting %ds...", attempt, max_retries, retry_after)
            time.sleep(retry_after)
            continue
 
        if resp.status_code in (200, 201, 204):
            logger.info("Inserted: %s -- %s [%s]",
                        metadata["meeting_name"], metadata["meeting_date"], metadata["meeting_type"])
            return True
 
        logger.warning("Attempt %d/%d: insert failed: %d %s",
                       attempt, max_retries, resp.status_code, resp.text[:300])
        if attempt < max_retries and resp.status_code in (500, 502, 503, 504):
            time.sleep(min(2 ** attempt, 30))
            continue
        return False
    return False
 
 
def _patch_record(dv_token, row_id, fields, max_retries=5):
    url = DATAVERSE_URL + "/api/data/v9.2/" + DATAVERSE_TABLE + "(" + row_id + ")"
    headers = _dv_headers(dv_token, write=True)
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.patch(url, headers=headers, json=fields, timeout=60)
        except requests.exceptions.RequestException as e:
            logger.warning("PATCH network error (%d/%d): %s", attempt, max_retries, e)
            time.sleep(min(2 ** attempt, 30))
            continue
        if resp.status_code == 429:
            time.sleep(int(resp.headers.get("Retry-After", 5)))
            continue
        if resp.status_code in (200, 204):
            return True
        if attempt < max_retries and resp.status_code in (500, 502, 503, 504):
            time.sleep(min(2 ** attempt, 30))
            continue
        logger.warning("PATCH failed: %d %s", resp.status_code, resp.text[:200])
        return False
    return False
 
 
def _delete_record(dv_token, row_id, max_retries=5):
    url = DATAVERSE_URL + "/api/data/v9.2/" + DATAVERSE_TABLE + "(" + row_id + ")"
    headers = _dv_headers(dv_token, write=True)
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.delete(url, headers=headers, timeout=30)
        except requests.exceptions.RequestException:
            time.sleep(min(2 ** attempt, 30))
            continue
        if resp.status_code == 429:
            time.sleep(int(resp.headers.get("Retry-After", 5)))
            continue
        if resp.status_code in (200, 204):
            return True
        if attempt < max_retries and resp.status_code in (500, 502, 503, 504):
            time.sleep(min(2 ** attempt, 30))
            continue
        logger.warning("DELETE failed: %d %s", resp.status_code, resp.text[:200])
        return False
    return False
 

# ---------------------------------------------------------------------------
# SharePoint -> Dataverse delete sync ("prune").
# When a .docx is deleted in SharePoint, its Dataverse row(s) are deleted on
# the next run. SharePoint is the source of truth. Guardrails:
#   * only rows inside the scanned root's year/month are considered, so a
#     --root override (e.g. Transcripts/2026/June) can never touch other months
#   * refuses to run if the SharePoint listing came back empty
#   * refuses to delete more than PRUNE_MAX_DELETE_FRACTION of in-scope rows
#     (when 10+ rows would go) -- protects against a bad/partial listing
# Disable with PRUNE_ORPHANS=false. Insert/update logic is untouched.
PRUNE_ORPHANS = os.environ.get("PRUNE_ORPHANS", "true").strip().lower() in ("1", "true", "yes")
PRUNE_MAX_DELETE_FRACTION = float(os.environ.get("PRUNE_MAX_DELETE_FRACTION", "0.20"))


def _prune_scope_from_root(root):
    """'Transcripts/2026' -> (2026, None); 'Transcripts/2026/June' -> (2026, 'June').
    Returns (None, None) when no year folder is present (prune is then skipped)."""
    parts = [p for p in (root or "").strip("/").split("/") if p]
    for i, p in enumerate(parts):
        if re.fullmatch(r"20\d{2}", p):
            month = parts[i + 1] if i + 1 < len(parts) else None
            return int(p), month
    return None, None


def _row_in_prune_scope(row, year, month):
    try:
        row_year = int(float(row.get("crb6b_year")))
    except (TypeError, ValueError):
        return False  # cannot verify scope -> leave the row alone
    if row_year != year:
        return False
    if month and (row.get("crb6b_month") or "").strip().lower() != month.strip().lower():
        return False
    return True


async def prune_index(root=None, dry_run=False, files=None):
    """Delete Dataverse rows whose SharePoint .docx no longer exists.
    Pass files= to reuse an already-fetched SharePoint listing."""
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- cannot prune")
        return

    root = root or TRANSCRIPTS_ROOT
    year, month = _prune_scope_from_root(root)
    print("=" * 60)
    print("PRUNE: " + root + ("  (DRY RUN -- nothing deleted)" if dry_run else ""))
    print("=" * 60)

    if year is None:
        print("Root has no year folder (e.g. Transcripts/2026) -- cannot scope safely, skipping prune.")
        return

    if files is None:
        graph_token = _get_graph_token()
        files = _list_docx_files(graph_token, root)  # raises on partial listing -- intentional
    sp_keys = {_norm_filename(f["name"]) for f in files}
    print("SharePoint files       : %d" % len(files))

    if not sp_keys:
        print("SharePoint listing is EMPTY -- refusing to prune (safety stop).")
        return

    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(
        dv_token,
        select="crb6b_table1id,crb6b_fileurl,crb6b_meetingdate,crb6b_meetingtype,crb6b_year,crb6b_month")
    in_scope = [r for r in rows if _row_in_prune_scope(r, year, month)]
    print("Dataverse rows in scope: %d of %d (%s%s)"
          % (len(in_scope), len(rows), year, "/" + month if month else ""))

    orphans, blank = [], 0
    for r in in_scope:
        key = _norm_filename(r.get("crb6b_fileurl", ""))
        if not key:
            blank += 1  # unidentifiable row -- never delete what we cannot match
            continue
        if key not in sp_keys:
            orphans.append(r)

    print("Orphan rows (file deleted in SharePoint): %d" % len(orphans))
    if blank:
        print("Rows with blank file URL (left untouched): %d" % blank)

    if not orphans:
        print("Nothing to prune.")
        print("=" * 60)
        return

    for r in orphans:
        print("   * %s | %s | %s"
              % (_norm_filename(r.get("crb6b_fileurl", ""))[:60],
                 str(r.get("crb6b_meetingdate") or "")[:10],
                 r.get("crb6b_meetingtype") or ""))

    if not dry_run and len(orphans) >= 10:
        fraction = len(orphans) / float(len(in_scope))
        if fraction > PRUNE_MAX_DELETE_FRACTION:
            print("\nSAFETY STOP: would delete %d of %d in-scope rows (%.0f%% > %.0f%% limit)."
                  % (len(orphans), len(in_scope), fraction * 100, PRUNE_MAX_DELETE_FRACTION * 100))
            print("Review with --prune --dry-run; if intentional, raise PRUNE_MAX_DELETE_FRACTION.")
            print("=" * 60)
            return

    deleted = failed = 0
    work_count = 0
    for r in orphans:
        if dry_run:
            deleted += 1
            continue
        if work_count and work_count % 80 == 0:
            dv_token = _get_dataverse_token()
            print("  (refreshed Dataverse token after %d deletes)" % work_count)
        work_count += 1
        if _delete_record(dv_token, r.get("crb6b_table1id")):
            deleted += 1
        else:
            failed += 1

    print("\n" + "=" * 60)
    print(("DRY RUN would delete : %d row(s)" if dry_run else "Deleted              : %d row(s)") % deleted)
    if failed:
        print("Failed               : %d row(s) (will retry next run)" % failed)
    print("=" * 60)


async def dedupe_index(dry_run=False):
    """Collapse duplicate rows: keep ONE row per file (the one with the longest
    summary), delete the rest. Pass dry_run=True to preview without deleting."""
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- cannot dedupe")
        return
    print("=" * 60)
    print("DEDUPE" + ("  (DRY RUN — nothing deleted)" if dry_run else ""))
    print("=" * 60)
 
    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(
        dv_token,
        select="crb6b_table1id,crb6b_fileurl,crb6b_meetingdate,crb6b_meetingtype,crb6b_summary")
 
    groups = defaultdict(list)
    for r in rows:
        key = _norm_filename(r.get("crb6b_fileurl", ""))
        if not key:
            key = "(blank)|%s|%s" % (str(r.get("crb6b_meetingdate"))[:10], r.get("crb6b_meetingtype"))
        groups[key].append(r)
 
    dup_groups = {k: v for k, v in groups.items() if len(v) > 1}
    extra = sum(len(v) - 1 for v in dup_groups.values())
    print("Total rows        : %d" % len(rows))
    print("Unique files      : %d" % len(groups))
    print("Duplicate groups  : %d" % len(dup_groups))
    print("Extra rows to drop: %d\n" % extra)
 
    deleted = 0
    for key, grp in dup_groups.items():
        grp_sorted = sorted(grp, key=lambda r: len((r.get("crb6b_summary") or "")), reverse=True)
        for r in grp_sorted[1:]:
            rid = r.get("crb6b_table1id")
            if dry_run:
                deleted += 1
                continue
            if _delete_record(dv_token, rid):
                deleted += 1
 
    print("=" * 60)
    print(("DRY RUN would delete: %d row(s)" if dry_run else "Deleted: %d row(s)") % deleted)
    print("Rows remaining: %d" % (len(rows) - (0 if dry_run else deleted)))
    print("=" * 60)
 
 
async def retype_index(root=None):
    """Re-derive meeting_type + meeting_name from each row's stored file URL using
    the current MEETING_TYPES table, and PATCH the row in place. Does NOT touch
    crb6b_summary, crb6b_fileurl, or crb6b_meetingdate. Fast (no LLM calls)."""
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- cannot retype")
        return
    print("=" * 60)
    print("RETYPE: re-derive type + name on existing rows")
    print("=" * 60)
 
    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(
        dv_token,
        select="crb6b_table1id,crb6b_fileurl,crb6b_meetingname,crb6b_meetingtype")
    print("Dataverse rows: %d\n" % len(rows))
 
    patched = unchanged = skipped = failed = 0
    work_count = 0
    for r in rows:
        row_id = r.get("crb6b_table1id")
        file_url = r.get("crb6b_fileurl") or ""
        fname = _norm_filename(file_url)
        if not fname:
            skipped += 1
            continue
 
        md = _parse_filename(fname)
        if not md:
            skipped += 1
            continue
 
        new_type = md["meeting_type"]
        new_name = md["meeting_name"]
        if (r.get("crb6b_meetingtype") == new_type and
                r.get("crb6b_meetingname") == new_name):
            unchanged += 1
            continue
 
        if work_count and work_count % 80 == 0:
            dv_token = _get_dataverse_token()
            print("  (refreshed Dataverse token after %d patches)" % work_count)
        work_count += 1
 
        if _patch_record(dv_token, row_id, {
                "crb6b_meetingtype": new_type,
                "crb6b_meetingname": new_name}):
            patched += 1
            print("RETYPED: %s  ->  %s | %s" % (fname[:50], new_type, new_name))
        else:
            failed += 1
 
    print("\n" + "=" * 60)
    print("RETYPE REPORT")
    print("  Patched   : %d" % patched)
    print("  Unchanged : %d" % unchanged)
    print("  Skipped   : %d  (blank file URL or unparseable filename)" % skipped)
    print("  Failed    : %d" % failed)
    print("=" * 60)
 
 
async def generate_meeting_index(root=None):
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- skipping meeting index generation")
        return
 
    root = root or TRANSCRIPTS_ROOT
    print("=" * 60)
    print("SHAREPOINT SCAN: " + root)
    print("=" * 60)
 
    graph_token = _get_graph_token()
    files = _list_docx_files(graph_token, root)
    print("\nFound %d .docx files\n" % len(files))
 
    dv_token = _get_dataverse_token()
    existing = _existing_filename_set(_get_existing_rows(dv_token))
    print("Existing records in Dataverse: %d\n" % len(existing))
 
    inserted = skipped = unparsed = failed = unnamed_inserted = 0
    unparsed_files, failed_files = [], []
    work_count = 0
 
    for f in files:
        filename = f["name"]
        metadata = _parse_filename(filename, fallback_dt=f.get("created") or f.get("modified"))
        if not metadata:
            print("UNPARSED (no date anywhere): " + filename)
            unparsed += 1
            unparsed_files.append(filename)
            continue
 
        key = _norm_filename(filename)
        if key in existing:
            skipped += 1
            continue
 
        if work_count and work_count % 40 == 0:
            graph_token = _get_graph_token()
            dv_token = _get_dataverse_token()
            print("  (refreshed auth tokens after %d inserts)" % work_count)
        work_count += 1
 
        tag = "INSERTING (no name detected): " if metadata["unnamed"] else "INSERTING: "
        print(tag + filename)
        print("  -> " + metadata["meeting_name"] + " | " + metadata["meeting_date"]
              + " | " + metadata["meeting_type"])
 
        text = _fetch_transcript_text(graph_token, f.get("id"), max_chars=1000000)
        digest = _make_digest(text)
 
        if not digest:
            # NEVER insert a row without a summary. A blank row is treated as
            # "already existing" on every later run and stays blank forever
            # (only --backfill would ever touch it again). Leaving the file out
            # means the retry pass below / the next run tries again instead.
            failed += 1
            failed_files.append((f, metadata))
            print("   NO DIGEST (download or LLM failed) -- row NOT inserted, will retry\n")
            continue
 
        if _upsert_record(dv_token, metadata, f["webUrl"], digest, text):
            existing.add(key)
            inserted += 1
            unnamed_inserted += 1 if metadata["unnamed"] else 0
            print("   SUCCESS\n")
        else:
            failed += 1
            failed_files.append((f, metadata))
            print("   FAILED (will retry)\n")
 
    if failed_files:
        print("Retrying %d failed insert(s)...\n" % len(failed_files))
        graph_token = _get_graph_token()
        dv_token = _get_dataverse_token()
        still = []
        # Re-read what is REALLY in Dataverse first: a first-pass insert that
        # timed out client-side may still have committed server-side, and
        # re-inserting it here is exactly how duplicate rows get created.
        existing = _existing_filename_set(_get_existing_rows(dv_token))
        for f, metadata in failed_files:
            key = _norm_filename(f["name"])
            if key in existing:
                print("ALREADY LANDED (first attempt committed): " + f["name"])
                inserted += 1
                failed -= 1
                continue
            text = _fetch_transcript_text(graph_token, f.get("id"))
            digest = _make_digest(text)
            if not digest:
                still.append(f["name"])
                continue
            if _upsert_record(dv_token, metadata, f["webUrl"], digest, text, max_retries=6):
                existing.add(key)
                inserted += 1
                failed -= 1
            else:
                still.append(f["name"])
        failed_files = still
    else:
        failed_files = []
 
    print("=" * 60)
    print("SYNC REPORT")
    print("=" * 60)
    print("Total SharePoint files : %d" % len(files))
    print("Already in Dataverse   : %d" % skipped)
    print("Newly inserted         : %d" % inserted)
    print("  (of which unnamed)   : %d" % unnamed_inserted)
    print("Failed                 : %d" % failed)
    print("Unparsed               : %d" % unparsed)
    print("=" * 60)
 
    if unparsed_files:
        print("\nUNPARSED FILES:")
        for u in unparsed_files:
            print("   * " + u)
    if failed_files:
        print("\nSTILL FAILED TO INSERT:")
        for u in failed_files:
            print("   * " + u)
 
    if unparsed == 0 and failed == 0:
        print("\nCLEAN: every SharePoint file is now represented in Dataverse.")
    else:
        print("\nWARNING: %d unparsed + %d failed file(s) are NOT in Dataverse." % (unparsed, failed))
 
    if PRUNE_ORPHANS:
        try:
            await prune_index(root=root, files=files)
        except Exception as e:
            logger.warning("Prune step failed (sync results above are unaffected): %s", e)
 
    return {"total": len(files), "inserted": inserted, "skipped": skipped,
            "unparsed": unparsed, "failed": failed}
 
 
async def backfill_index(root=None, force=False):
    """Fill crb6b_summary on existing rows (no re-insert)."""
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- cannot backfill")
        return
 
    root = root or TRANSCRIPTS_ROOT
    print("=" * 60)
    print("BACKFILL: " + root + ("  (force=overwrite all)" if force else "  (empty summaries only)"))
    print("=" * 60)
 
    graph_token = _get_graph_token()
    files = _list_docx_files(graph_token, root)
    id_by_name = {_norm_filename(f["name"]): f["id"] for f in files}
    print("SharePoint files mapped: %d" % len(id_by_name))
 
    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(dv_token, select="crb6b_table1id,crb6b_fileurl,crb6b_summary")
    print("Dataverse rows: %d\n" % len(rows))
 
    patched = skipped = orphan = failed = 0
    work_count = 0
    for r in rows:
        row_id = r.get("crb6b_table1id")
        if (r.get("crb6b_summary") or "").strip() and not force:
            skipped += 1
            continue
        item_id = id_by_name.get(_norm_filename(r.get("crb6b_fileurl", "")))
        if not item_id:
            orphan += 1
            print("ORPHAN (no SP file): " + (r.get("crb6b_fileurl") or "")[:80])
            continue
        if work_count and work_count % 40 == 0:
            graph_token = _get_graph_token()
            dv_token = _get_dataverse_token()
            print("  (refreshed auth tokens after %d rows)" % work_count)
        work_count += 1
        try:
            digest = _make_digest(_fetch_transcript_text(graph_token, item_id))
            if not digest:
                # Do NOT patch "" over the row and call it FILLED -- that is
                # how a failing backfill silently cements blanks in place.
                failed += 1
                print("NO DIGEST (download or LLM failed) -- left as-is: "
                      + (r.get("crb6b_fileurl") or "")[:60])
                continue
            if _patch_record(dv_token, row_id, {"crb6b_summary": digest}):
                patched += 1
                print("FILLED: " + (r.get("crb6b_fileurl") or "")[:60])
            else:
                failed += 1
        except Exception as e:
            failed += 1
            logger.warning("Row %s failed, skipping: %s", row_id, e)
 
    print("\n" + "=" * 60)
    print("BACKFILL REPORT")
    print("  Patched (filled)  : %d" % patched)
    print("  Skipped (had text): %d" % skipped)
    print("  Orphan rows       : %d" % orphan)
    print("  Failed            : %d" % failed)
    print("=" * 60)
    print("Orphan rows = stray/duplicate rows whose file URL has no SharePoint file.")
 
 
async def verify_index(root=None):
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- cannot verify")
        return
 
    root = root or TRANSCRIPTS_ROOT
    print("=" * 60)
    print("VERIFY MODE: " + root)
    print("=" * 60)
 
    graph_token = _get_graph_token()
    files = _list_docx_files(graph_token, root)
    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(dv_token)
    existing = _existing_filename_set(rows)
 
    print("\nSharePoint files found : %d" % len(files))
    print("Dataverse records found: %d\n" % len(existing))
 
    missing, unparsed = [], []
    for f in files:
        metadata = _parse_filename(f["name"], fallback_dt=f.get("created") or f.get("modified"))
        if not metadata:
            unparsed.append(f["name"])
            continue
        if _norm_filename(f["name"]) not in existing:
            missing.append((f["name"], metadata))
 
    print("Missing from Dataverse : %d" % len(missing))
    print("Unparsed filenames     : %d" % len(unparsed))
 
    if missing:
        print("\nMISSING FILES:")
        for fn, md in missing:
            print("   * " + fn + "  ->  " + md["meeting_name"] + " | "
                  + md["meeting_date"] + " | " + md["meeting_type"])
    else:
        print("\nAll SharePoint files are present in Dataverse!")
    if unparsed:
        print("\nUNPARSED FILES:")
        for u in unparsed:
            print("   * " + u)
    print("=" * 60)
 
 
async def reconcile_index(root=None):
    if not DATAVERSE_URL:
        logger.warning("DATAVERSE_URL not set -- cannot reconcile")
        return
 
    root = root or TRANSCRIPTS_ROOT
    print("=" * 70)
    print("RECONCILE: " + root)
    print("=" * 70)
 
    graph_token = _get_graph_token()
    files = _list_docx_files(graph_token, root)
    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(dv_token)
 
    sp = defaultdict(lambda: defaultdict(int))
    for f in files:
        md = _parse_filename(f["name"], fallback_dt=f.get("created") or f.get("modified"))
        if md:
            sp[md["meeting_type"]][md["meeting_date"]] += 1
 
    dv = defaultdict(lambda: defaultdict(int))
    for r in rows:
        t = r.get("crb6b_meetingtype") or "Other"
        d = str(r.get("crb6b_meetingdate") or "")[:10]
        if d:
            dv[t][d] += 1
 
    all_types = sorted(set(sp) | set(dv))
    for t in all_types:
        dates = sorted(set(sp[t]) | set(dv[t]))
        print("\n%s" % t)
        print("  %-12s %-10s %-10s %s" % ("DATE", "SHAREPOINT", "DATAVERSE", "STATUS"))
        for d in dates:
            s, v = sp[t].get(d, 0), dv[t].get(d, 0)
            if s and not v:
                status = "MISSING in Dataverse"
            elif v and not s:
                status = "ORPHAN (in DV, no SP file)"
            elif s != v:
                status = "COUNT MISMATCH"
            else:
                status = "ok"
            print("  %-12s %-10d %-10d %s" % (d, s, v, status))
 
    print("\n" + "=" * 70)
    print("MISSING = sync hasn't captured it (or the .docx isn't in SharePoint yet).")
    print("ORPHAN / COUNT MISMATCH = duplicate or stale rows to clean up.")
    print("=" * 70)
 
 
SELFTEST_CASES = [
    ("20260608_111650_Daily_Operational_Focus.docx", None, "2026-06-08", "DOF"),
    ("2026-06-08_111650_Daily_Operational_Focus.docx", None, "2026-06-08", "DOF"),
    ("2026-06-04_SAM_125536.docx", None, "2026-06-04", "SAM"),
    ("20260611_070000UTC_Daily_Operational_Focus.docx", None, "2026-06-11", "DOF"),
    ("Daily_Operational_Focus_20260611_070000UTC.docx", None, "2026-06-11", "DOF"),
    ("HUSAM-AutoRecord-20260512_063000.docx", None, "2026-05-12", "HUSAM"),
    ("2026-06-12_DOF_080000.docx", None, "2026-06-12", "DOF"),
    ("OCC Alignment Meeting-Meeting Recording-2026-06-15.docx", None, "2026-06-15", "OCC"),
    ("Weekend Call_2026-06-13.docx", None, "2026-06-13", "Weekend Call"),
    ("2026-06-16_dof_080000.DOCX", None, "2026-06-16", "DOF"),
    ("EOD_2026-06-12.docx", None, "2026-06-12", "EOD"),
    ("Leadership Team Alignment_20260610_090000.docx", None, "2026-06-10", "Leadership Team"),
    ("garbage_no_date.docx", datetime(2026, 6, 11), "2026-06-11", "Other"),
]
 
 
def run_selftest():
    failures = 0
    for fn, fb, exp_date, exp_type in SELFTEST_CASES:
        md = _parse_filename(fn, fallback_dt=fb)
        if md is None:
            print("FAIL  %s -> returned None" % fn)
            failures += 1
            continue
        if md["meeting_date"] != exp_date or md["meeting_type"] != exp_type:
            print("FAIL  %s -> date=%s type=%s (expected %s / %s)"
                  % (fn, md["meeting_date"], md["meeting_type"], exp_date, exp_type))
            failures += 1
        else:
            print("ok    %s -> %s | %s | %s"
                  % (fn, md["meeting_name"], md["meeting_date"], md["meeting_type"]))
 
    for fn, fb, _, _ in SELFTEST_CASES:
        assert _parse_filename(fn, fallback_dt=fb) is not None, "dropped: " + fn
 
    a = _norm_filename("2026-06-12_DOF_080000.docx")
    b = _norm_filename("2026-06-12_DOF_153000.docx")
    if a == b:
        print("FAIL  same-day distinct recordings collide")
        failures += 1
    else:
        print("ok    same-day distinct recordings get distinct keys")
 
    u = ("https://stpgweb.sharepoint.com/sites/ChatSTP/Shared%20Documents/"
         "Transcripts/2026/June/2026-06-12_DOF_080000.docx?web=1")
    if _norm_filename(u) != _norm_filename("2026-06-12_DOF_080000.docx"):
        print("FAIL  URL/filename normalization mismatch")
        failures += 1
    else:
        print("ok    URL and filename normalize to the same dedup key")
 
    layouts = ("https://stpgweb.sharepoint.com/sites/ChatSTP/_layouts/15/Doc.aspx?"
               "sourcedoc=%7B0B548376-5777-43E1-9B18-1EE1C960B702%7D&"
               "file=2026-06-18_DOF_111753.docx&action=default&mobileredirect=true")
    if _norm_filename(layouts) != _norm_filename("2026-06-18_DOF_111753.docx"):
        print("FAIL  _layouts URL did not resolve to its ?file= filename")
        failures += 1
    else:
        print("ok    _layouts/Doc.aspx URL resolves to the ?file= filename")
 
    dof_variants = ["DOF", "dof", "Daily_Operational_Focus", "DAILY OPERATIONAL FOCUS"]
    canon = {_parse_filename("2026-06-12_%s.docx" % v)["meeting_type"] for v in dof_variants}
    if canon != {"DOF"}:
        print("FAIL  DOF variants not canonical: %s" % canon)
        failures += 1
    else:
        print("ok    all DOF spellings canonicalize to 'DOF'")
 
    scope_cases = [
        ("Transcripts/2026", (2026, None)),
        ("Transcripts/2026/June", (2026, "June")),
        ("Transcripts", (None, None)),
    ]
    for sc_root, sc_exp in scope_cases:
        got = _prune_scope_from_root(sc_root)
        if got != sc_exp:
            print("FAIL  prune scope %s -> %s (expected %s)" % (sc_root, got, sc_exp))
            failures += 1
        else:
            print("ok    prune scope %-22s -> %s" % (sc_root, str(got)))
 
    print("\n%s" % ("SELFTEST PASSED" if failures == 0 else "SELFTEST FAILED (%d)" % failures))
    return failures == 0
 
 
if __name__ == "__main__":
    import asyncio
 
    parser = argparse.ArgumentParser(description="Sync/verify SharePoint transcripts -> Dataverse index")
    parser.add_argument("--verify", action="store_true", help="Read-only gap check")
    parser.add_argument("--reconcile", action="store_true", help="Per-type, per-date coverage report")
    parser.add_argument("--backfill", action="store_true", help="Fill crb6b_summary on existing rows")
    parser.add_argument("--force", action="store_true", help="With --backfill: overwrite existing summaries")
    parser.add_argument("--dedupe", action="store_true", help="Delete duplicate rows (keep one per file)")
    parser.add_argument("--prune", action="store_true", help="Delete Dataverse rows whose SharePoint file no longer exists")
    parser.add_argument("--dry-run", action="store_true", help="With --dedupe/--prune: preview only, delete nothing")
    parser.add_argument("--retype", action="store_true", help="Re-derive type+name on existing rows from MEETING_TYPES (no re-digest)")
    parser.add_argument("--selftest", action="store_true", help="Offline logic tests (no tenant)")
    parser.add_argument("--root", default=None, help="Override TRANSCRIPTS_ROOT (e.g. Transcripts/2026)")
    args = parser.parse_args()
 
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S")
 
    if args.selftest:
        sys.exit(0 if run_selftest() else 1)
    elif args.dedupe:
        asyncio.run(dedupe_index(dry_run=args.dry_run))
    elif args.prune:
        asyncio.run(prune_index(root=args.root, dry_run=args.dry_run))
    elif args.retype:
        asyncio.run(retype_index(root=args.root))
    elif args.reconcile:
        asyncio.run(reconcile_index(root=args.root))
    elif args.backfill:
        asyncio.run(backfill_index(root=args.root, force=args.force))
    elif args.verify:
        asyncio.run(verify_index(root=args.root))
    else:
        asyncio.run(generate_meeting_index(root=args.root))