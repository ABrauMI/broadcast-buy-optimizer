# Deploying the Sample Buy Slack app to Railway

This runs the bot as a Railway "worker" -- a background process with no
public URL, which is all a Socket Mode Slack app needs (it opens an
outbound connection to Slack; nothing needs to reach it from the internet).

Do the Slack-side setup in `SLACK_SETUP.md` first if you haven't -- you need
the `SLACK_BOT_TOKEN` and `SLACK_APP_TOKEN` values from that before this
will actually connect.

## 1. Create the project

1. In Railway, **New Project** -> **Deploy from GitHub repo**.
2. Pick `ABrauMI/broadcast-buy-optimizer`. If Railway doesn't see it yet,
   it'll prompt you to install/authorize the Railway GitHub App for that repo.
3. Branch: `main` (this is also the repo's default branch, so Railway
   should pick it automatically).

Railway will detect the `requirements.txt` and build the project with
Nixpacks automatically -- no Dockerfile needed.

## 2. Set the start command

The repo includes a `Procfile` (`worker: python3 -m slack_app.app`), but
Railway is most reliable when you also set this explicitly:

1. Open the service -> **Settings** -> **Deploy**.
2. **Custom Start Command**: `python3 -m slack_app.app`

## 3. Add the environment variables

Service -> **Variables** -> **New Variable**, add both:

- `SLACK_BOT_TOKEN` = `xoxb-...` (from Slack's OAuth & Permissions page)
- `SLACK_APP_TOKEN` = `xapp-...` (from Slack's Socket Mode page)

## 4. Skip the public domain

Under **Settings** -> **Networking**, do **not** click "Generate Domain."
This service doesn't accept incoming traffic -- it only makes an outbound
connection to Slack, so a public URL isn't needed and would just be unused.

## 5. Deploy

Railway deploys automatically once the repo, start command, and variables
are set. Check the **Deployments** tab -> **View Logs** -- you should see
Bolt's Socket Mode connection log lines with no errors. If the token is
wrong or a scope is missing, the logs will show exactly that.

From here, `/sample-buy` and `/strata-order` in Slack should work the same
as running it locally, except it now stays up whether or not your laptop is
open. Railway also restarts the process automatically if it ever crashes.

## Updating later

Any time you push new commits to `main` (or whichever branch the service is
tracking), Railway redeploys automatically.
