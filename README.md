# The Second Seat

Sign-up page for the second Seahawks season seat at Lumen Field.

- `index.html` is the whole page. GitHub Pages serves it.
- `apps-script/Code.gs` is the backend. It lives inside the Google Sheet and keeps the
  list in the `Games` and `Requests` tabs, and the seat location in `Settings`.
- `dev/` is a local test rig. It runs `Code.gs` in the browser against a fake spreadsheet.

No names, codes, or seat numbers are stored in this repo. Friends reach the list through a
link that ends in a crew code; host actions need a separate host code. Both live in the
script's properties. The seat location lives in the sheet and is only sent to someone who
already has a code.

Keep the Google Sheet's sharing on **Restricted**. Anyone who can edit the sheet can open
the script and read both codes.

## Switching sign-ups on (one time)

1. Open the Google Sheet, then **Extensions > Apps Script**.
2. Replace everything in `Code.gs` with the contents of `apps-script/Code.gs`. Save.
3. Pick `setUp` in the function dropdown and press **Run**. Approve the permission prompt
   (it only asks for this one spreadsheet). This adds the `Games`, `Requests` and
   `Settings` tabs, carries over names and the seat location from the old sign-up tab,
   and makes the two codes.
4. **Deploy > New deployment > Web app**. Execute as **Me**, who has access **Anyone**. Deploy.
5. Copy the web app URL (it ends in `/exec`) into `API_URL` near the top of the script in
   `index.html`, and push.
6. Back in the Sheet, reload, then **Second Seat > Show links**. Send friends the first link.
   Open the second one yourself once on each device you host from.
7. Set the Sheet's sharing back to **Restricted**. Friends no longer need to edit it.

## Running it day to day

- Give a seat, mark it paid, release it, or remove a name from the page in host mode.
- Change a kickoff time or a price in the `Games` tab, or the seat location in `Settings`.
  The page picks it up within a minute.
- To switch everything off: **Deploy > Manage deployments > Archive**.
- After editing `Code.gs` again: **Deploy > Manage deployments > edit > New version**, so the
  URL stays the same.
- **Second Seat > New codes** replaces both codes. Old links stop working.

## Testing locally

    python dev/serve.py

Then open <http://localhost:8335/>. `/plain` shows the page exactly as it ships.
