import azure.functions as func
import logging
import asyncio
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

app = func.FunctionApp()

CENTRAL = ZoneInfo("America/Chicago")

_WEEKDAY_SLOTS = {(7, 30), (9, 30), (11, 30), (13, 30), (15, 30), (17, 30), (19, 30)}
_WEEKEND_SLOTS = {(9, 0), (19, 0)}
_GRACE_MIN = 12  


def _is_due(now_ct: datetime) -> bool:
    m = now_ct.minute
    if 0 <= m <= _GRACE_MIN:
        slot = (now_ct.hour, 0)
    elif 30 <= m <= 30 + _GRACE_MIN:
        slot = (now_ct.hour, 30)
    else:
        return False
    # Python weekday(): Mon=0 .. Sun=6. >= 4 = Fri, Sat, Sun.
    is_weekend = now_ct.weekday() >= 4
    return slot in (_WEEKEND_SLOTS if is_weekend else _WEEKDAY_SLOTS)


@app.timer_trigger(
    schedule="0 0,30 * * * *",
    arg_name="mytimer",
    run_on_startup=False,
    use_monitor=True,
)
def transcript_extractor(mytimer: func.TimerRequest) -> None:
    """Timer trigger: pull transcripts -> SharePoint, then sync the Dataverse meeting index."""
    now_ct = datetime.now(CENTRAL)
    if not _is_due(now_ct):
        return  # not a target slot, wake up and go back to sleep cheaply

    logging.info("ChatSTP Transcript Extractor triggered (%s CT)", now_ct.strftime("%a %H:%M"))

    if mytimer.past_due:
        logging.warning("Timer is past due - running anyway")

    from transcript_extractor.transcript_extractor_v37 import _setup_logging

    _setup_logging()
    asyncio.run(_run_all())

    logging.info("ChatSTP Transcript Extractor completed")


async def _run_all() -> None:
    """Run the transcript pull, then sync the Dataverse meeting index off the new docx."""
    # 1) Pull transcripts and upload .docx to SharePoint (unchanged).
    from transcript_extractor.transcript_extractor_v37 import run_job
    await run_job()

    try:
        from transcript_extractor.meeting_index_v2 import generate_meeting_index
        root = f"Transcripts/{datetime.now(timezone.utc).year}"
        logging.info("Meeting index sync starting (root=%s)", root)
        summary = await generate_meeting_index(root=root)
        logging.info("Meeting index sync done: %s", summary)
    except Exception:
        # Don't fail the whole run if indexing hiccups -- transcripts are already uploaded.
        logging.exception("Meeting index sync failed (transcripts were still uploaded)")