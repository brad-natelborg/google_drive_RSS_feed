const PODCAST = Object.freeze({
  title: "Brad's Tailored Podcasts",
  description: 'Useful, interesting audio made to fit the moment—one tailored episode at a time.',
  author: "Brad's Tailored Podcasts",
  language: 'en-us',
  category: 'Society & Culture',
  subcategory: 'Personal Journals',
});

const HEADERS = Object.freeze([
  'Episode',
  'Title',
  'Description',
  'Published',
  'Duration',
  'Audio File ID',
  'Audio Bytes',
  'GUID',
  'Published?',
]);

function doGet(e) {
  const format = String((e && e.parameter && e.parameter.format) || 'rss').toLowerCase();
  if (format === 'json') {
    return ContentService.createTextOutput(JSON.stringify(buildPodcastData_(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(buildRss_())
    .setMimeType(ContentService.MimeType.RSS);
}

function setupPodcast() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('EPISODE_SHEET_ID')) {
    throw new Error('Podcast storage is already configured for this script.');
  }

  const folder = DriveApp.createFolder(PODCAST.title);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const spreadsheet = SpreadsheetApp.create(`${PODCAST.title} - Episodes`);
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName('Episodes');
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange('D:D').setNumberFormat('ddd, mmm d, yyyy h:mm AM/PM');
  sheet.autoResizeColumns(1, HEADERS.length);

  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  spreadsheetFile.moveTo(folder);

  properties.setProperties({
    EPISODE_SHEET_ID: spreadsheet.getId(),
    PODCAST_FOLDER_ID: folder.getId(),
  });

  Logger.log(`Folder: ${folder.getUrl()}`);
  Logger.log(`Episode sheet: ${spreadsheet.getUrl()}`);
  return { folderUrl: folder.getUrl(), sheetUrl: spreadsheet.getUrl() };
}

function setCoverArt(fileId) {
  const file = DriveApp.getFileById(String(fileId).trim());
  makePublic_(file);
  PropertiesService.getScriptProperties().setProperty('COVER_FILE_ID', file.getId());
  return file.getUrl();
}

function addEpisode(title, description, audioFileId, duration) {
  const sheet = getEpisodeSheet_();
  const episode = Math.max(0, ...readEpisodes_().map(item => item.number)) + 1;
  const file = DriveApp.getFileById(String(audioFileId).trim());
  makePublic_(file);
  const published = new Date();
  const guid = `brads-tailored-podcasts-episode-${String(episode).padStart(3, '0')}`;
  sheet.appendRow([
    episode,
    title,
    description,
    published,
    normalizeDuration_(duration),
    file.getId(),
    file.getSize(),
    guid,
    true,
  ]);
  sheet
  .getRange(sheet.getLastRow(), 9)
  .insertCheckboxes()
  .setValue(true);
  return { episode, guid, fileUrl: file.getUrl() };
}

function refreshAudioMetadata() {
  const sheet = getEpisodeSheet_();
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row += 1) {
    const fileId = String(values[row][5] || '').trim();
    if (!fileId) continue;
    const file = DriveApp.getFileById(fileId);
    makePublic_(file);
    sheet.getRange(row + 1, 7).setValue(file.getSize());
  }
}

function buildPodcastData_() {
  const properties = PropertiesService.getScriptProperties();
  const coverFileId = properties.getProperty('COVER_FILE_ID');
  if (!coverFileId) throw new Error('Cover art is not configured. Run setCoverArt(fileId).');
  const cover = DriveApp.getFileById(coverFileId);
  makePublic_(cover);
  const episodes = readEpisodes_().filter(item => item.published).sort((a, b) => b.number - a.number);
  return {
    ...PODCAST,
    coverUrl: directDownloadUrl_(cover),
    episodes,
  };
}

function buildRss_() {
  const data = buildPodcastData_();
  const webAppUrl = ScriptApp.getService().getUrl();
  const latestDate = data.episodes.length ? data.episodes[0].publishedDate : new Date();
  const items = data.episodes.map(episode => `    <item>
      <title>${xml_(episode.title)}</title>
      <description>${xml_(episode.description)}</description>
      <content:encoded><![CDATA[<p>${cdata_(episode.description)}</p>]]></content:encoded>
      <guid isPermaLink="false">${xml_(episode.guid)}</guid>
      <pubDate>${episode.publishedDate.toUTCString()}</pubDate>
      <enclosure url="${xml_(episode.audioUrl)}" length="${episode.audioBytes}" type="${xml_(episode.mimeType)}" />
      <itunes:episode>${episode.number}</itunes:episode>
      <itunes:episodeType>full</itunes:episodeType>
      <itunes:duration>${xml_(episode.duration)}</itunes:duration>
      <itunes:explicit>false</itunes:explicit>
      <itunes:image href="${xml_(data.coverUrl)}" />
    </item>`).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0"
  xmlns:atom="http://www.w3.org/2005/Atom"
  xmlns:content="http://purl.org/rss/1.0/modules/content/"
  xmlns:itunes="http://www.itunes.com/dtds/podcast-1.0.dtd">
  <channel>
    <title>${xml_(data.title)}</title>
    <link>${xml_(webAppUrl)}</link>
    <description>${xml_(data.description)}</description>
    <language>${xml_(data.language)}</language>
    <lastBuildDate>${latestDate.toUTCString()}</lastBuildDate>
    <atom:link href="${xml_(webAppUrl)}" rel="self" type="application/rss+xml" />
    <itunes:author>${xml_(data.author)}</itunes:author>
    <itunes:summary>${xml_(data.description)}</itunes:summary>
    <itunes:type>episodic</itunes:type>
    <itunes:explicit>false</itunes:explicit>
    <itunes:image href="${xml_(data.coverUrl)}" />
    <itunes:category text="${xml_(data.category)}"><itunes:category text="${xml_(data.subcategory)}" /></itunes:category>
    <image><url>${xml_(data.coverUrl)}</url><title>${xml_(data.title)}</title><link>${xml_(webAppUrl)}</link></image>
${items}
  </channel>
</rss>`;
}

function readEpisodes_() {
  const values = getEpisodeSheet_().getDataRange().getValues();
  return values.slice(1).filter(row => Number(row[0]) > 0 && String(row[5]).trim()).map(row => {
    const file = DriveApp.getFileById(String(row[5]).trim());
    const published = row[3] instanceof Date ? row[3] : new Date(row[3]);
    return {
      number: Number(row[0]),
      title: String(row[1]),
      description: String(row[2]),
      publishedDate: published,
      duration: normalizeDuration_(row[4]),
      audioUrl: directDownloadUrl_(file),
      audioBytes: Number(row[6]) || file.getSize(),
      guid: String(row[7]) || `brads-tailored-podcasts-episode-${String(row[0]).padStart(3, '0')}`,
      published: row[8] === true,
      mimeType: podcastMimeType_(file),
    };
  });
}

function getEpisodeSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('EPISODE_SHEET_ID');
  if (!id) throw new Error('Podcast storage is not configured. Run setupPodcast() first.');
  return SpreadsheetApp.openById(id).getSheetByName('Episodes');
}

function directDownloadUrl_(file) {
  const resourceKey = file.getResourceKey();
  const key = resourceKey ? `&resourcekey=${encodeURIComponent(resourceKey)}` : '';
  return `https://drive.usercontent.google.com/download?id=${encodeURIComponent(file.getId())}&export=download&confirm=t${key}`;
}

function podcastMimeType_(file) {
  const type = String(file.getMimeType() || '').toLowerCase();
  if (type === 'audio/mp4' || /\.m4a$/i.test(file.getName())) return 'audio/mp4';
  return type.startsWith('audio/') ? type : 'audio/mpeg';
}

function makePublic_(file) {
  if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
}

function normalizeDuration_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'UTC', 'HH:mm:ss');
  }
  const text = String(value || '').trim();
  if (/^\d{1,2}:\d{2}:\d{2}$/.test(text)) return text.padStart(8, '0');
  if (/^\d{1,3}:\d{2}$/.test(text)) return `00:${text}`;
  throw new Error(`Duration must be HH:MM:SS or MM:SS; received: ${text}`);
}

function xml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function cdata_(value) {
  return String(value).replace(/]]>/g, ']]]]><![CDATA[>');
}

function configureInitialPodcast() {
  setCoverArt("1i_ffMj_I9R_kc7D2WvePEy92-8NVXe4k");

  addEpisode(
    "Turning TSI Sensors into Predictive AI Networks",
    "How connected TSI sensor data can become the foundation for predictive AI networks, enabling earlier detection, smarter decisions, and proactive operational insight.",
    "1LooPCidLnMh4hWESmaYCjVr1DcZOEPPE",
    "00:48:06"
  );
}

function syncPodcastFolder() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const properties = PropertiesService.getScriptProperties();
    const folderId = properties.getProperty("PODCAST_FOLDER_ID");

    if (!folderId) {
      throw new Error("Podcast folder is not configured.");
    }

    const folder = DriveApp.getFolderById(folderId);
    const sheet = getEpisodeSheet_();
    const episodes = readEpisodes_();

    const knownFileIds = new Set(
      episodes.map(function (episode) {
        return episode.audioFileId;
      })
    );

    const existingValues = sheet.getDataRange().getValues();
    existingValues.slice(1).forEach(function (row) {
      if (row[5]) knownFileIds.add(String(row[5]).trim());
    });

    let nextEpisode = Math.max(
      0,
      ...existingValues.slice(1).map(function (row) {
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

      if (!isAudio || knownFileIds.has(file.getId())) {
        continue;
      }

      makePublic_(file);

      const title = fileName
        .replace(/\.[^.]+$/, "")
        .replace(/_/g, " ")
        .trim();

      const description =
        file.getDescription() ||
        "A new episode of Brad's Tailored Podcasts.";

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
        true
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

function repairEpisodeSheetRows() {
  const sheet = getEpisodeSheet_();
  const currentRows = sheet.getDataRange().getValues();

  const episodes = currentRows.slice(1).filter(function (row) {
    return Number(row[0]) > 0 && String(row[5] || "").trim();
  });

  sheet.clear();

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange("D:D").setNumberFormat("ddd, mmm d, yyyy h:mm AM/PM");

  if (episodes.length) {
    sheet
      .getRange(2, 1, episodes.length, HEADERS.length)
      .setValues(episodes);

    sheet
      .getRange(2, 9, episodes.length, 1)
      .insertCheckboxes()
      .setValues(
        episodes.map(function (row) {
          return [row[8] === true];
        })
      );
  }

  const desiredRows = Math.max(episodes.length + 1, 2);
  const excessRows = sheet.getMaxRows() - desiredRows;

  if (excessRows > 0) {
    sheet.deleteRows(desiredRows + 1, excessRows);
  }

  sheet.autoResizeColumns(1, HEADERS.length);
}