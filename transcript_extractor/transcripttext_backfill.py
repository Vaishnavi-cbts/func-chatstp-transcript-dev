"""
transcripttext_backfill.py — fill crb6b_transcripttext on all existing rows.

For every Dataverse row that has no transcript text yet:
  1. finds its .docx in SharePoint (GUID match, filename fallback)
  2. downloads it and extracts the clean transcript text
     (drops the duplicated "Raw Transcript (VTT)" dump, keeps up to 1M chars)
  3. PATCHes crb6b_transcripttext on the row

Rows that already have text are SKIPPED, so it is safe to re-run after a
failure or after new meetings are synced. --force re-fills everything.

DRY RUN by default (shows the plan + one real extraction sample).
RUN (same folder as meeting_index_v2.py + duplicate_cleanup_v2.py):
    python transcripttext_backfill.py
    python transcripttext_backfill.py --execute
"""

import argparse
import re
import time

from meeting_index_v2 import (
    DATAVERSE_URL,
    DATAVERSE_TABLE,
    TRANSCRIPTS_ROOT,
    GRAPH_BASE,
    DRIVE_ID,
    _get_graph_token,
    _get_dataverse_token,
    _get_existing_rows,
    _graph_get_bytes,
    _extract_docx_text,
    _patch_record,
)
from duplicate_cleanup_v2 import list_sp_files

GUID_RE = re.compile(r"sourcedoc=(?:%7B|\{)([0-9A-Fa-f\-]+)(?:%7D|\})", re.IGNORECASE)


def guid_of(url):
    m = GUID_RE.search(url or "")
    return m.group(1).upper() if m else None


def fname_of(url):
    """Filename fallback for rows whose fileurl contains a real path."""
    tail = (url or "").rstrip("/").split("/")[-1]
    return tail if tail.lower().endswith(".docx") else ""


def main():
    ap = argparse.ArgumentParser(description="Fill crb6b_transcripttext on existing rows")
    ap.add_argument("--execute", action="store_true", help="Apply (default: DRY RUN)")
    ap.add_argument("--force", action="store_true", help="Re-fill rows that already have text")
    args = ap.parse_args()
    dry_run = not args.execute

    print("=" * 74)
    print("TRANSCRIPT TEXT BACKFILL" + ("  (DRY RUN - nothing changes)" if dry_run
                                  else "  *** EXECUTE MODE ***"))
    print("=" * 74)

    print("\n[1/3] Listing SharePoint files...")
    graph_token = _get_graph_token()
    files = list_sp_files(graph_token, TRANSCRIPTS_ROOT)
    by_guid = {}
    by_name = {}
    for f in files:
        g = guid_of(f.get("webUrl", ""))
        if g:
            by_guid[g] = f
        by_name[f["name"]] = f
    print("      %d files (%d with GUID links)" % (len(files), len(by_guid)))

    print("[2/3] Reading Dataverse rows...")
    dv_token = _get_dataverse_token()
    rows = _get_existing_rows(
        dv_token,
        select="crb6b_table1id,crb6b_fileurl,crb6b_transcripttext,crb6b_meetingdate,crb6b_meetingtype")
    print("      rows: %d\n" % len(rows))

    plan, already, orphan = [], 0, 0
    for r in rows:
        if (r.get("crb6b_transcripttext") or "").strip() and not args.force:
            already += 1
            continue
        item = by_guid.get(guid_of(r.get("crb6b_fileurl", "")))
        if not item:
            item = by_name.get(fname_of(r.get("crb6b_fileurl", "")), None)
        if not item or not item.get("id"):
            orphan += 1
            print("  ORPHAN (no SP file): %s" % (r.get("crb6b_fileurl") or "")[:70])
            continue
        plan.append((r, item))

    print("[3/3] Plan")
    print("-" * 74)
    print("  to fill          : %d" % len(plan))
    print("  already have text: %d (skipped)" % already)
    print("  orphan rows      : %d (no matching file - left alone)" % orphan)

    if not plan:
        print("\nNothing to do.")
        return

    # one real extraction sample so you can verify BEFORE the full run
    r0, it0 = plan[0]
    sample = _extract_docx_text(
        _graph_get_bytes(GRAPH_BASE + "/drives/" + DRIVE_ID + "/items/" + it0["id"] + "/content",
                         graph_token),
        max_chars=1000000)
    print("\nSample extraction from: %s" % it0["name"])
    print("  extracted %d characters" % len(sample))
    print("  first line: %s" % (sample.splitlines() or [""])[0][:90])

    if dry_run:
        print("\nDRY RUN complete. If the sample looks like real transcript text, re-run with --execute.")
        return

    patched = failed = 0
    for i, (r, item) in enumerate(plan, 1):
        if i % 40 == 0:
            graph_token = _get_graph_token()
            dv_token = _get_dataverse_token()
            print("  (refreshed tokens at row %d)" % i)
        try:
            text = _extract_docx_text(
                _graph_get_bytes(GRAPH_BASE + "/drives/" + DRIVE_ID + "/items/" + item["id"] + "/content",
                                 graph_token),
                max_chars=1000000)
            if not text.strip():
                failed += 1
                print("  EMPTY TEXT (left as-is): %s" % item["name"][:70])
                continue
            if _patch_record(dv_token, r["crb6b_table1id"], {"crb6b_transcripttext": text}):
                patched += 1
            else:
                failed += 1
        except Exception as e:
            failed += 1
            print("  FAILED %s: %s" % (item["name"][:60], e))
        if i % 50 == 0:
            print("  ... %d/%d done (%d filled, %d failed)" % (i, len(plan), patched, failed))

    print("\n" + "=" * 74)
    print("BACKFILL REPORT: filled %d | failed %d | skipped (had text) %d | orphan %d"
          % (patched, failed, already, orphan))
    print("=" * 74)


if __name__ == "__main__":
    main()
