# Running the Sample Buy Slack app

The Slack side has to be created through Slack's own dashboard -- that part
can't be done for you. Everything below maps directly to what `slack_app/app.py`
expects, so following it exactly will make the bot work the first time.

## 1. Create the app

1. Go to https://api.slack.com/apps -> **Create New App** -> **From scratch**.
2. Name it (e.g. "Sample Buy") and pick your workspace.

## 2. Turn on Socket Mode

1. Left sidebar -> **Socket Mode** -> toggle it on.
2. It'll prompt you to create an App-Level Token -- name it anything (e.g.
   `socket-token`), scope `connections:write` is added automatically.
3. Copy the generated token (starts with `xapp-`) -- this is `SLACK_APP_TOKEN`.

## 3. Add bot permissions

Left sidebar -> **OAuth & Permissions** -> **Scopes** -> **Bot Token Scopes**.
Add all of these:

- `commands`
- `chat:write`
- `files:read`
- `files:write`
- `channels:history`
- `groups:history`
- `im:history`
- `mpim:history`

## 4. Create the slash commands

Left sidebar -> **Slash Commands** -> **Create New Command**, twice:

- Command: `/sample-buy`
  Short description: `Build a sample TV buy from rate cards`
- Command: `/strata-order`
  Short description: `Generate a Strata order from a Sample Buy workbook`

For both, Request URL isn't used with Socket Mode -- put anything, e.g.
`https://example.com`.

## 5. Turn on Interactivity

Left sidebar -> **Interactivity & Shortcuts** -> toggle on. Request URL isn't
used here either with Socket Mode.

## 6. Subscribe to message events

Left sidebar -> **Event Subscriptions** -> toggle on -> under **Subscribe to
bot events**, add:

- `message.channels`
- `message.groups`
- `message.im`
- `message.mpim`

(This is what lets the bot notice the rate card files you reply with in a thread.)

## 7. Install the app

**OAuth & Permissions** -> **Install to Workspace** -> approve. Copy the
**Bot User OAuth Token** (starts with `xoxb-`) -- this is `SLACK_BOT_TOKEN`.

## 8. Invite the bot

In whichever Slack channel you want to use it, run `/invite @Sample Buy`
(or whatever you named the app).

## 9. Run it

```bash
pip install -r requirements.txt
export SLACK_BOT_TOKEN=xoxb-...
export SLACK_APP_TOKEN=xapp-...
python3 -m slack_app.app
```

Leave that process running (Socket Mode keeps an open connection to Slack --
no public URL or hosting needed to try it out, though for regular team use
you'll eventually want it running somewhere persistent, like a small always-on
VM, rather than a laptop).

## Using it

### Building the initial workbook: `/sample-buy`

1. In the channel, type `/sample-buy`.
2. Fill in the modal: target GRPs, demo, and the earliest/latest spot time.
   Optionally also fill in market name, flight start/end date, and campaign
   name -- leave all four blank to skip this, or fill in all of market/start/end
   to also get a Strata-importable order built straight from the freshly
   computed buy.
3. The bot posts a message with a **Build Sample Buy** button. Reply to that
   message (in the thread) with your station rate card XML files -- one
   message or several, any number of stations.
4. Click **Build Sample Buy**. The bot runs the same logic as the CLI and
   uploads the resulting workbook (and, if you filled in the Strata fields,
   a `.sbx` order file) right into the thread.

### Reviewing/editing first, then generating the order: `/strata-order`

For when a buyer wants to open the Sample Buy workbook, make changes (day
quantities, rates, added or removed rows), and only then generate the
Strata order from what they actually approved:

1. Download the `.xlsx` from the `/sample-buy` thread and edit it as needed
   in Excel -- only the Daypart, Station, Program, Time, Length, and the
   day columns (M/T/W/Th/F/Sa/Su) and Rate/Rating matter; everything else
   in the sheet is just live formulas for review.
2. In the channel, type `/strata-order`. Fill in market name and flight
   start/end date (required this time, since generating the order is the
   whole point) and, optionally, a campaign name.
3. The bot posts a message with a **Generate Strata Order** button. Reply
   to that message with the (possibly edited) `.xlsx` file.
4. Click **Generate Strata Order**. The bot reads the workbook exactly as
   it currently stands and uploads the resulting `.sbx`.
