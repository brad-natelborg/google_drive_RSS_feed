# Google Drive Podcast RSS Feed

Host a small personal podcast using Google Drive, Google Sheets, and Google Apps Script.

The project creates:

- A public RSS 2.0 podcast feed served by an Apps Script web app
- A Google Drive folder for cover art and episode audio
- A Google Sheet that acts as the episode catalog
- Optional automatic discovery of audio files dropped into the Drive folder

This project is intended for personal and low-volume podcast feeds. Google Drive is not a dedicated podcast CDN, so it may not be suitable for large public audiences or high download volumes.

## How it works

```text
Google Drive audio files
          ↓
Google Sheets episode catalog
          ↓
Google Apps Script web app
          ↓
RSS feed → podcast application
```

Apps Script generates the RSS document each time the feed is requested. Episode audio and cover art are downloaded from files shared through Google Drive.

## Requirements

- A Google account
- Permission to use Google Apps Script, Drive, and Sheets
- The ability to share Drive files as **Anyone with the link**
- Podcast audio in MP3, M4A, AAC, WAV, OGG, or Opus format
- Square JPG or PNG cover art

Google Workspace administrators can disable anonymous Drive sharing or public Apps Script deployments. If the relevant sharing options are unavailable, contact the Workspace administrator or use a personal Google account.

## Repository files

```text
Code.gs            Main Apps Script code
appsscript.json     Apps Script manifest and OAuth scopes
README.md           Setup and usage instructions
```

## 1. Create the Apps Script project

1. Go to [script.google.com](https://script.google.com/).
2. Select **New project**.
3. Rename the project, for example `Brad's Tailored Podcasts RSS`.
4. Replace the contents of `Code.gs` with this repository's `Code.gs`.
5. Open **Project Settings**.
6. Enable **Show "appsscript.json" manifest file in editor**.
7. Return to **Editor**, open `appsscript.json`, and replace its contents with the repository's `appsscript.json`.
8. Select **Save project**.

Before saving, change the `timeZone` in `appsscript.json` if necessary.

## 2. Customize the podcast

Edit the `PODCAST` object near the top of `Code.gs`:

```javascript
const PODCAST = Object.freeze({
  title: "Brad's Tailored Podcasts",
  description: "Useful, interesting audio made to fit the moment—one tailored episode at a time.",
  author: "Brad's Tailored Podcasts",
  language: "en-us",
  category: "Society & Culture",
  subcategory: "Personal Journals",
});
```

Save the project after making changes.

## 3. Create the Drive folder and episode sheet

1. Open the function menu above the editor.
2. Select `setupPodcast`.
3. Select **Run**.
4. Select **Review permissions**.
5. Choose the Google account that will own the podcast.
6. Approve access to Google Drive and Google Sheets.

If Google displays an unverified-app warning for your own script, select **Advanced**, continue to the project, and approve the requested permissions.

The function creates:

- A Drive folder named after the podcast
- A spreadsheet named `<Podcast title> - Episodes`

Run `setupPodcast` only once. The folder and sheet IDs are saved in the project's Script Properties.

## 4. Upload and configure the cover art

1. Open the newly created podcast folder in Google Drive.
2. Upload the square cover image.
3. Open the file's sharing link.
4. Copy the file ID from the URL:

```text
https://drive.google.com/file/d/FILE_ID/view
```

Add a temporary helper function to the bottom of `Code.gs`:

```javascript
function configureCover() {
  setCoverArt("FILE_ID");
}
```

Replace `FILE_ID`, save, select `configureCover`, and select **Run**. The script makes the cover publicly viewable and saves its ID.

## 5. Add an episode manually

Upload the episode audio to the podcast Drive folder and copy its Drive file ID. Then add or edit this helper function:

```javascript
function addMyNewEpisode() {
  addEpisode(
    "Episode title",
    "A short description of the episode.",
    "AUDIO_FILE_ID",
    "00:15:30"
  );
}
```

The duration must use `HH:MM:SS` or `MM:SS` format.

Save the project, select `addMyNewEpisode`, and select **Run**. Run the helper once per episode. The script automatically:

- Assigns the next episode number
- Makes the audio publicly viewable
- Records the file size and MIME type
- Generates a permanent GUID
- Adds the episode to the spreadsheet
- Marks the episode as published

The spreadsheet is the live episode database. Titles, descriptions, dates, durations, and publication status can be edited directly in the sheet.

## 6. Optional: scan the Drive folder automatically

Add these functions to `Code.gs` to discover new audio files every five minutes:

```javascript
function syncPodcastFolder() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const folderId = PropertiesService.getScriptProperties()
      .getProperty("PODCAST_FOLDER_ID");

    if (!folderId) {
      throw new Error("Podcast folder is not configured.");
    }

    const folder = DriveApp.getFolderById(folderId);
    const sheet = getEpisodeSheet_();
    const rows = sheet.getDataRange().getValues().slice(1);

    const knownFileIds = new Set(
      rows
        .map(function (row) {
          return String(row[5] || "").trim();
        })
        .filter(Boolean)
    );

    let nextEpisode = Math.max(
      0,
      ...rows.map(function (row) {
        return Number(row[0]) || 0;
      })
    ) + 1;

    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = String(file.getMimeType()).toLowerCase();

      const isAudio =
        mimeType.startsWith("audio/") ||
        /\.(mp3|m4a|aac|wav|ogg|opus)$/i.test(fileName);

      if (!isAudio || knownFileIds.has(file.getId())) continue;

      makePublic_(file);

      const title = fileName
        .replace(/\.[^.]+$/, "")
        .replace(/_/g, " ")
        .trim();

      const description =
        file.getDescription() ||
        "A new episode of " + PODCAST.title + ".";

      const guid =
        "brads-tailored-podcasts-episode-" +
        String(nextEpisode).padStart(3, "0");

      sheet.appendRow([
        nextEpisode,
        title,
        description,
        file.getDateCreated(),
        "00:00:00",
        file.getId(),
        file.getSize(),
        guid,
        true,
      ]);

      sheet
        .getRange(sheet.getLastRow(), 9)
        .insertCheckboxes()
        .setValue(true);

      knownFileIds.add(file.getId());
      nextEpisode += 1;
    }
  } finally {
    lock.releaseLock();
  }
}

function installPodcastFolderScanner() {
  ScriptApp.getProjectTriggers()
    .filter(function (trigger) {
      return trigger.getHandlerFunction() === "syncPodcastFolder";
    })
    .forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

  ScriptApp.newTrigger("syncPodcastFolder")
    .timeBased()
    .everyMinutes(5)
    .create();

  syncPodcastFolder();
}
```

Save the project, select `installPodcastFolderScanner`, and select **Run**. New audio files dropped into the podcast folder will be added within approximately five minutes.

Use the desired episode title as the filename. For example:

```text
Turning TSI Sensors into Predictive AI Networks.m4a
```

Apps Script cannot reliably determine media duration from a Drive file. Automatically discovered episodes initially use `00:00:00`; enter the exact duration in the spreadsheet afterward.

## 7. Deploy the RSS feed

1. Select **Deploy** and then **New deployment**.
2. Select the gear icon beside **Select type**.
3. Choose **Web app**.
4. Set **Execute as** to **Me**.
5. Set **Who has access** to **Anyone**.
6. Select **Deploy**.
7. Approve permissions if requested.
8. Copy the web app URL ending in `/exec`.

The `/exec` URL is the podcast RSS address to enter in a podcast application.

To inspect the catalog as JSON, open:

```text
YOUR_WEB_APP_URL?format=json
```

## 8. Test the feed

Verify the following before subscribing:

1. Open the `/exec` URL in a private or incognito browser window.
2. Confirm it loads without requiring a Google login.
3. Open `/exec?format=json` and confirm every published episode appears.
4. Open each returned audio URL and confirm it downloads without authentication.
5. Add the plain `/exec` URL to the podcast application.

If the JSON endpoint shows an episode but the podcast app does not, confirm that every episode has a unique GUID and episode number. Podcast applications may also cache RSS responses; refreshing the feed or temporarily adding a harmless query parameter such as `?v=2` can force a new subscription request.

## Updating the code

Spreadsheet edits are read live and do not require a new deployment.

After changing `Code.gs`, update the deployed version:

1. Select **Deploy** and then **Manage deployments**.
2. Select the pencil icon for the web app.
3. Select **New version**.
4. Select **Deploy**.

The existing `/exec` URL remains unchanged.

## Spreadsheet columns

| Column | Purpose |
|---|---|
| Episode | Permanent episode number |
| Title | Episode title |
| Description | Episode summary or show notes |
| Published | Publication date and time |
| Duration | `HH:MM:SS` playback duration |
| Audio File ID | Google Drive file ID |
| Audio Bytes | File size used by the RSS enclosure |
| GUID | Permanent unique episode identifier |
| Published? | Checked episodes are included in the RSS feed |

Do not reuse a GUID or episode number. Podcast applications use GUIDs to distinguish episodes.

## Troubleshooting

### The spreadsheet is blank

`setupPodcast` creates an empty catalog. Upload an audio file and run `addEpisode` through a helper function, or install the automatic folder scanner.

### Episodes begin around row 1,001

An older version of the setup code added checkbox validation to the entire `I` column, causing Sheets to consider all 1,000 rows used. Remove this line if present:

```javascript
sheet.getRange("I:I").insertCheckboxes();
```

Move the episode rows directly below the header or recreate the sheet with the corrected setup code.

### The JSON endpoint shows fewer episodes than the spreadsheet

Confirm that each episode row has:

- A positive episode number
- An Audio File ID
- A valid publication date
- A unique GUID
- The **Published?** checkbox checked

### The feed shows two items but the podcast app shows one

Ensure the two rows have different episode numbers and GUID values. Duplicate GUIDs are treated as the same episode.

### A Drive audio link requests sign-in

Open the audio file's Share dialog and set General access to **Anyone with the link** and Viewer. Workspace administrators may prohibit this setting.

### Large files do not play reliably

Google Drive may show a confirmation page for unusually large files or high traffic. Convert the recording to a reasonably sized MP3, or use dedicated object storage or a podcast host for the audio while continuing to use Apps Script for the RSS document.

### The feed did not change after editing code

Apps Script web app deployments use versions. Update the deployment through **Manage deployments** and select **New version**.

## Security and privacy

- The web app runs as the account that deploys it.
- The RSS endpoint is public.
- Cover art and published audio files are shared publicly by link.
- Do not store private recordings in the public podcast folder.
- Keep the spreadsheet and Apps Script project restricted to trusted editors.

## Limitations

- Google Drive is not a podcast CDN.
- Apps Script cannot reliably extract audio duration.
- Drive download behavior and public-sharing availability can vary by account policy.
- High download volumes can trigger Google quotas or Drive safeguards.
- This project does not submit the feed to podcast directories automatically.

## License

Add the license appropriate for your repository before distributing or accepting external contributions.
