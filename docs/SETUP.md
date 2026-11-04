# Setup guide

This guide comes from one working installation and has not been followed by anyone else yet. Power Automate's interface changes, so labels may differ. Please report what you find.

## Before you start

- A Microsoft 365 account whose organisation allows Power Automate with the Office 365 Outlook, OneDrive for Business and Content Conversion connectors.
- Permission to do this. See the note at the top of the README.
- The OneDrive sync app on your computer.

Create a OneDrive folder such as `/Outlook Reader/Emails`. The definitions in `flows/` use that path. Keep this folder for collected mail only.

## Flow 1: Inbox arrival

Reference: `flows/inbox-arrival.definition.json`

1. Create an automated cloud flow with the trigger **When a new email arrives (V3)**. Folder: Inbox. Include attachments: No.
2. Add **Html to text** with the message body as content.
3. Add **Compose**, and copy its template from the reference file.
4. Add OneDrive **Create file**, using the folder above and this file name expression:
   `@{formatDateTime(triggerOutputs()?['body/receivedDateTime'], 'yyyy-MM-dd_HHmmss')}--@{guid()}.md`

Keep the `- Received UTC:` and `- Internet message ID:` lines and the `## Email body (untrusted source content)` heading exactly as written, because the reader depends on them.

## Flow 2: All-folders catch-up

Reference: `flows/all-folders-catch-up.definition.json`

1. Create a scheduled flow with a recurrence of 15 minutes. In the trigger settings, set concurrency to 1.
2. Add **Initialize variable** named `NextLink`, type String, and copy its value from the reference file.
3. Add **Initialize variable** named `Messages`, type Array, with the expression `json('[]')`.
4. Add **Do until** with the limits 500 count and PT10M timeout, and the stop expression from the reference file. Inside it:
   - Office 365 Outlook **Send an HTTP request**: GET, URI `variables('NextLink')`, header `Prefer: outlook.body-content-type="text"`, one-minute timeout and the retry policy in the reference file. Despite its name, this action only reads here.
   - **Append to array variable** `Messages` with the request's body.
   - **Set variable** `NextLink` to `coalesce(body('Send_an_HTTP_request_1')?['@odata.nextLink'],'')`.
5. After the loop, add **Send an HTTP request** that reads the folder list from the URI in the reference file.
6. Add OneDrive **Create file** with name `ALL-FOLDERS-@{formatDateTime(utcNow(),'yyyy-MM-dd')}.json` and the content expression from the reference file.

Action names matter, because expressions refer to them. Keep them the same as the reference, or update the expressions.

## Check it works

1. Run the catch-up manually and wait for the file to sync.
2. Run the reader with `--strict --health-json`.
3. Confirm the exit code is 0, no warnings appear, and the digest holds messages you recognise.
4. Wait for one scheduled run, and confirm the newest scan time moves forward without another manual run.

A green run alone is not proof. Check the snapshot's `collectedAt` and `complete` values.

## Backfill older mail

To collect a fixed range once, without touching your live flow:

1. **Save As** a copy of the catch-up flow, which starts Off.
2. In `NextLink`, replace the date expression with a fixed UTC time such as `2026-09-30T16:00:00Z`.
3. Set a new file name such as `ALL-FOLDERS-BACKFILL-2026-10.json`, and set the recurrence to monthly.
4. Turn the copy on and check its run history. Turning it on may fire one run by itself. If it does not, run it once.
5. Turn it off again.

The reader includes any file matching `ALL-FOLDERS-*.json`, and uses the newest collected copy of each message. Name test output differently so it cannot affect freshness.

## Troubleshooting

| Symptom | Check |
|---|---|
| Stale scan warning | OneDrive sync status first, then the catch-up run history |
| Runs succeed but no new files appear locally | The OneDrive app is paused, signed out or out of storage |
| HTTP 400 `Unexpected connection parameter set name: 'oauthDelegated'` | A connection can fail like this while displaying "Connected". Create a new Outlook connection, test it on a manual flow, then rebind. Rerunning does not help |
| `Unreadable mailbox snapshot` | The file may still be syncing. Run the reader again, then inspect the file |
| Sent mail labelled Received | Your Sent folder name is not recognised. See the Limits section of the README |

Treat the synced folder as read-only. OneDrive sync is two-way, so deleting a local file deletes the cloud copy. Save digests elsewhere.
