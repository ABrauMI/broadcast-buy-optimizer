"""GPS Impact Sample Buy Slack app.

Flow:
  1. /sample-buy opens a modal for the buy parameters (GRP target, demo,
     daypart window).
  2. Submitting the modal posts a message in the channel with those
     parameters and a "Build Sample Buy" button, and opens a session keyed
     by that message's ts.
  3. Reply to that message with your station rate card XML files attached
     (one file per message or several -- any number of replies is fine).
     The bot confirms each one it picks up.
  4. Click "Build Sample Buy" and the bot runs the existing pipeline and
     uploads the resulting workbook back into the thread.

Run with:
  SLACK_BOT_TOKEN=xoxb-... SLACK_APP_TOKEN=xapp-... python3 -m slack_app.app
"""

import json
import logging
import os
import tempfile

import requests
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

from broadcast_buy.pipeline import run_pipeline
from .session import sessions

BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("sample_buy_slack_app")

app = App(token=BOT_TOKEN)

DEFAULTS = {
    "target_grps": "750",
    "demo_group": "Adults",
    "demo_age": "35",
    "earliest_time": "7:00",
    "latest_time": "23:00",
}


def _modal_view(channel_id):
    def text_input(block_id, label, initial):
        return {
            "type": "input",
            "block_id": block_id,
            "label": {"type": "plain_text", "text": label},
            "element": {
                "type": "plain_text_input",
                "action_id": "value",
                "initial_value": initial,
            },
        }

    return {
        "type": "modal",
        "callback_id": "sample_buy_modal",
        "private_metadata": json.dumps({"channel": channel_id}),
        "title": {"type": "plain_text", "text": "New Sample Buy"},
        "submit": {"type": "plain_text", "text": "Next: upload rate cards"},
        "close": {"type": "plain_text", "text": "Cancel"},
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": "Set the buy parameters, then you'll upload station rate card XML files in the next step.",
                },
            },
            text_input("target_grps", "Target weekly GRPs", DEFAULTS["target_grps"]),
            text_input("demo_group", "Target demo group (e.g. Adults)", DEFAULTS["demo_group"]),
            text_input("demo_age", "Target demo age (e.g. 35)", DEFAULTS["demo_age"]),
            text_input("earliest_time", "Earliest spot time (HH:MM)", DEFAULTS["earliest_time"]),
            text_input("latest_time", "Latest spot time (HH:MM)", DEFAULTS["latest_time"]),
        ],
    }


@app.command("/sample-buy")
def handle_sample_buy_command(ack, body, client):
    ack()
    client.views_open(trigger_id=body["trigger_id"], view=_modal_view(body["channel_id"]))


@app.view("sample_buy_modal")
def handle_modal_submission(ack, body, client, view):
    values = view["state"]["values"]
    raw = {k: (v["value"]["value"] or "").strip() for k, v in values.items()}

    errors = {}
    try:
        target_grps = float(raw["target_grps"] or DEFAULTS["target_grps"])
    except ValueError:
        errors["target_grps"] = "Enter a number, e.g. 750"
    try:
        demo_age = int(raw["demo_age"] or DEFAULTS["demo_age"])
    except ValueError:
        errors["demo_age"] = "Enter a whole number, e.g. 35"
    for field in ("earliest_time", "latest_time"):
        text = raw[field] or DEFAULTS[field]
        if ":" not in text or not text.replace(":", "").isdigit():
            errors[field] = "Use HH:MM, e.g. 7:00 or 23:00"

    if errors:
        ack(response_action="errors", errors=errors)
        return
    ack()

    metadata = json.loads(view["private_metadata"])
    channel = metadata["channel"]
    params = {
        "target_grps": target_grps,
        "demo_group": raw["demo_group"] or DEFAULTS["demo_group"],
        "demo_age": demo_age,
        "earliest_time": raw["earliest_time"] or DEFAULTS["earliest_time"],
        "latest_time": raw["latest_time"] or DEFAULTS["latest_time"],
    }

    user_id = body["user"]["id"]
    summary = (
        f"*New sample buy requested by <@{user_id}>*\n"
        f"Target: *{params['target_grps']:.0f} GRPs/week* | "
        f"Demo: *{params['demo_group']} {params['demo_age']}+* | "
        f"Window: *{params['earliest_time']}-{params['latest_time']}*\n\n"
        f"Reply in this thread with your station rate card XML files "
        f"(one message or several, any number of stations). "
        f"Click *Build Sample Buy* below once they're all in."
    )
    posted = client.chat_postMessage(
        channel=channel,
        text=summary,
        blocks=[
            {"type": "section", "text": {"type": "mrkdwn", "text": summary}},
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Build Sample Buy \U0001F4CA"},
                        "action_id": "build_sample_buy",
                        "style": "primary",
                    }
                ],
            },
        ],
    )
    sessions.create(posted["ts"], channel, params)


@app.event("message")
def handle_message_with_files(event, client, say):
    thread_ts = event.get("thread_ts")
    files = event.get("files")
    if not thread_ts or not files:
        return
    session = sessions.get(thread_ts)
    if session is None:
        return

    added = []
    for f in files:
        name = f.get("name", "rate_card.xml")
        if not name.lower().endswith(".xml"):
            continue
        url = f.get("url_private_download") or client.files_info(file=f["id"])["file"]["url_private_download"]
        r = requests.get(url, headers={"Authorization": f"Bearer {BOT_TOKEN}"})
        r.raise_for_status()
        tmp_path = os.path.join(tempfile.mkdtemp(prefix="sample_buy_"), name)
        with open(tmp_path, "wb") as out:
            out.write(r.content)
        session.add_file(tmp_path, name)
        added.append(name)

    if added:
        say(
            channel=session.channel,
            thread_ts=thread_ts,
            text=f"Got {', '.join(added)} -- {len(session.file_paths)} rate card(s) received so far.",
        )


@app.action("build_sample_buy")
def handle_build_button(ack, body, client):
    ack()
    thread_ts = body["message"]["ts"]
    channel = body["channel"]["id"]
    session = sessions.get(thread_ts)

    if session is None or not session.file_paths:
        client.chat_postMessage(
            channel=channel,
            thread_ts=thread_ts,
            text="I don't have any rate card XML files for this thread yet -- reply with them first, then click Build again.",
        )
        return

    client.chat_postMessage(
        channel=channel, thread_ts=thread_ts, text=f"Building the sample buy from {len(session.file_paths)} file(s)..."
    )

    output_dir = tempfile.mkdtemp(prefix="sample_buy_out_")
    output_path = os.path.join(output_dir, "sample_buy.xlsx")

    try:
        result, log_lines = run_pipeline(
            session.file_paths,
            output_path,
            target_grps=session.params["target_grps"],
            target_demo_group=session.params["demo_group"],
            target_demo_age=session.params["demo_age"],
            earliest_time=session.params["earliest_time"],
            latest_time=session.params["latest_time"],
        )
    except Exception:
        logger.exception("Sample buy pipeline failed")
        client.chat_postMessage(
            channel=channel,
            thread_ts=thread_ts,
            text="Something went wrong building the sample buy -- check the files uploaded are valid AAAA/TVB rate card XML.",
        )
        return

    blended_cpp = result.total_cost / result.achieved_grps if result.achieved_grps else 0
    summary = (
        f"*Sample buy complete*\n"
        f"Achieved: *{result.achieved_grps:.0f}* of *{session.params['target_grps']:.0f}* weekly GRPs\n"
        f"Weekly cost: *${result.total_cost:,.0f}*  |  Blended CPP: *${blended_cpp:,.0f}*  |  "
        f"Spots/week: *{len(result.spots)}*"
    )
    for w in result.warnings:
        summary += f"\n:warning: {w}"

    client.files_upload_v2(
        channel=channel,
        thread_ts=thread_ts,
        file=output_path,
        filename="sample_buy.xlsx",
        initial_comment=summary,
    )
    sessions.discard(thread_ts)


if __name__ == "__main__":
    handler = SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"])
    handler.start()
