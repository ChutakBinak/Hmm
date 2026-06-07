// ═══════════════════════════════════════════════════════════════════════════════
//  Hmm · Vocal Wellness — Google Apps Script Backend
//  ─────────────────────────────────────────────────
//  SETUP (one time):
//    1. Create a new Google Sheet and copy its ID from the URL.
//    2. Paste the ID into SPREADSHEET_ID below.
//    3. In Apps Script editor: Deploy → New deployment → Web App
//       • Execute as:  Me
//       • Who has access:  Anyone
//    4. Copy the Web App URL and paste it into APP_CONFIG.APPS_SCRIPT_URL in hmm.html
// ═══════════════════════════════════════════════════════════════════════════════

const SPREADSHEET_ID    = '1Ajfyfxf3X0Qc7DGnz_sA7RA4SQD_nJTfVy5hPDN_U8o';
const APP_VERSION       = '1.1';
const DRIVE_ROOT_FOLDER = 'Hmm_Recordings';   // folder created in your Drive

// ─── CORS / routing ───────────────────────────────────────────────────────────
function doGet(e) {
  return _json({ status: 'ok', app: 'Hmm Vocal Wellness', version: APP_VERSION });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const payload = JSON.parse(e.postData.contents);
    switch (payload.action) {
      case 'saveSession': return _json(saveSession(payload));
      case 'saveAudio':   return _json(saveAudioRecording(payload));
      default:            return _json({ error: 'Unknown action: ' + payload.action });
    }
  } catch (err) {
    return _json({ error: err.toString(), stack: err.stack });
  } finally {
    lock.releaseLock();
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Sheet bootstrapper ───────────────────────────────────────────────────────
function _sheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    // Style the header row
    const hdr = sh.getRange(1, 1, 1, headers.length);
    hdr.setBackground('#1a1a2e');
    hdr.setFontColor('#a78bfa');
    hdr.setFontWeight('bold');
    hdr.setFontFamily('Courier New');
    sh.setFrozenRows(1);
    // Auto-resize columns
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}

// ─── Main handler: save a full session ────────────────────────────────────────
function saveSession(payload) {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const { user, session, segments } = payload;

  if (!user || !user.sub) return { error: 'Missing user' };
  if (!session)           return { error: 'Missing session' };

  const now    = new Date();
  const nowISO = now.toISOString();

  // ── 1. Users ──────────────────────────────────────────────────────────────
  const usersSheet = _sheet(ss, 'Users', [
    'user_id', 'email', 'display_name', 'photo_url',
    'first_seen', 'last_seen', 'total_sessions',
  ]);

  const allRows    = usersSheet.getDataRange().getValues();
  const userRowIdx = allRows.findIndex((r, i) => i > 0 && r[0] === user.sub);

  if (userRowIdx === -1) {
    usersSheet.appendRow([
      user.sub, user.email, user.name, user.picture || '',
      nowISO, nowISO, 1,
    ]);
  } else {
    const row = userRowIdx + 1;
    usersSheet.getRange(row, 6).setValue(nowISO);
    usersSheet.getRange(row, 7).setValue((Number(allRows[userRowIdx][6]) || 0) + 1);
  }

  // ── 2. Sessions ───────────────────────────────────────────────────────────
  const sessSheet = _sheet(ss, 'Sessions', [
    // identity
    'session_id', 'user_id', 'email', 'display_name',
    // timing
    'timestamp_iso', 'date', 'time',
    'duration_ms', 'duration_str',
    // pitch
    'avg_freq_hz', 'peak_freq_hz',
    'pitch_stability_pct', 'pitch_accuracy_pct',
    // resonance
    'resonance_score', 'resonance_label',
    'harmonics_avg', 'hnr_pct',
    'amp_consistency_pct',
    // voice
    'voice_profile',
    // wellness
    'wellness_score',
    'score_stability_contrib', 'score_accuracy_contrib',
    'score_resonance_contrib', 'score_duration_contrib',
    // hum / pause
    'hum_count', 'pause_count', 'total_hum_ms', 'total_pause_ms',
    // session config
    'target_hz',
    // mood
    'pre_mood', 'post_mood',
    // volume / decibels
    'avg_db_fs', 'peak_db_fs', 'min_db_fs', 'db_range',
    // meta
    'notes', 'tags', 'program_id', 'app_version',
  ]);

  const d    = new Date(session.ts);
  const stab = session.stab  || 0;
  const acc  = session.acc   || 0;
  const res  = session.res   || 0;
  const ds   = Math.floor((session.dur || 0) / 1000);
  const durScore = Math.min(100, Math.round(ds / 300 * 100));

  sessSheet.appendRow([
    session.sessionId || `hmm_${session.ts}`,
    user.sub, user.email, user.name,
    d.toISOString(),
    d.toLocaleDateString('en-US', { year:'numeric', month:'2-digit', day:'2-digit' }),
    d.toLocaleTimeString([], { hour:'2-digit', minute:'2-digit' }),
    session.dur   || 0,
    session.durStr || '',
    session.avg   || 0,
    session.peak  || 0,
    stab, acc,
    res,
    session.resLbl || '',
    session.avgHarm  || 0,
    session.avgHNR   || 0,
    session.ampCons  || 0,
    session.voiceProfile || '',
    session.score || 0,
    Math.round(stab  * 0.32),
    Math.round(acc   * 0.22),
    Math.round(res   * 0.28),
    Math.round(durScore * 0.18),
    session.humCount    || 0,
    session.pauseCount  || 0,
    session.totalHumMs  || 0,
    session.totalPauseMs|| 0,
    session.tgt   || 0,
    session.preMood  || 0,
    session.postMood || 0,
    session.avgDb   != null ? session.avgDb   : '',
    session.peakDb  != null ? session.peakDb  : '',
    session.minDb   != null ? session.minDb   : '',
    session.dbRange != null ? session.dbRange : '',
    session.notes  || '',
    (session.tags || []).join(', '),
    session.programId || '',
    APP_VERSION,
  ]);

  // ── 3. Segments ───────────────────────────────────────────────────────────
  if (segments && segments.length > 0) {
    const segSheet = _sheet(ss, 'Segments', [
      'session_id', 'user_id', 'email',
      'seg_index', 'type',
      'start_ms', 'end_ms', 'duration_ms',
      'avg_frequency_hz', 'avg_resonance_pct', 'avg_stability_pct',
    ]);
    segments.forEach((seg, i) => {
      segSheet.appendRow([
        session.sessionId || `hmm_${session.ts}`,
        user.sub, user.email,
        i, seg.type || '',
        seg.startTime || 0,
        seg.endTime   || 0,
        seg.duration  || 0,
        seg.avgFrequency || '',
        seg.avgResonance || '',
        seg.avgStability || '',
      ]);
    });
  }

  return { success: true, sessionId: session.sessionId };
}

// ═══════════════════════════════════════════════════════════════════════════════
//  AUDIO RECORDING STORAGE
//  Saves audio blob to Google Drive in organized folder hierarchy:
//    Hmm_Recordings / {user_email} / {YYYY-MM} / {sessionId}.{ext}
//  Indexes every file in the Recordings sheet.
//  Links the audio URL back to the matching row in Sessions sheet.
// ═══════════════════════════════════════════════════════════════════════════════
function saveAudioRecording(payload) {
  const { user, sessionId, audioB64, mimeType, durationMs, sizeBytes } = payload;

  if (!user || !user.sub) return { error: 'Missing user' };
  if (!audioB64)           return { error: 'Missing audio data' };
  if (!sessionId)          return { error: 'Missing sessionId' };

  // ── 1. Resolve / create folder structure ──────────────────────────────────
  const rootFolder  = _driveFolder(null,       DRIVE_ROOT_FOLDER);
  const userFolder  = _driveFolder(rootFolder, _safeEmail(user.email));
  const d           = new Date();
  const monthStr    = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  const monthFolder = _driveFolder(userFolder, monthStr);

  // ── 2. Write audio file ───────────────────────────────────────────────────
  const mime   = mimeType || 'audio/webm';
  const ext    = mime.includes('mp4') ? 'mp4' : mime.includes('ogg') ? 'ogg' : 'webm';
  const fname  = `${sessionId}.${ext}`;
  const fpath  = `${DRIVE_ROOT_FOLDER}/${_safeEmail(user.email)}/${monthStr}/${fname}`;

  let decoded;
  try {
    decoded = Utilities.base64Decode(audioB64);
  } catch(e) {
    return { error: 'base64 decode failed: ' + e.toString() };
  }

  const blob = Utilities.newBlob(decoded, mime, fname);
  const file = monthFolder.createFile(blob);
  // Make viewable by anyone with the link (for admin analysis)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId  = file.getId();
  const fileUrl = `https://drive.google.com/file/d/${fileId}/view`;
  const now     = new Date().toISOString();

  // ── 3. Recordings sheet ───────────────────────────────────────────────────
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recSh   = _sheet(ss, 'Recordings', [
    'recording_id', 'session_id', 'user_id', 'email', 'display_name',
    'file_id',      'file_url',   'file_name', 'folder_path',
    'duration_ms',  'size_bytes', 'mime_type',
    'uploaded_at',  'month',      'app_version',
  ]);

  recSh.appendRow([
    `rec_${sessionId}`,
    sessionId,
    user.sub,
    user.email,
    user.name || '',
    fileId,
    fileUrl,
    fname,
    fpath,
    durationMs  || 0,
    sizeBytes   || 0,
    mime,
    now,
    monthStr,
    APP_VERSION,
  ]);

  // ── 4. Back-fill audio_url into Sessions sheet ────────────────────────────
  const sessSh = ss.getSheetByName('Sessions');
  if (sessSh) {
    const data    = sessSh.getDataRange().getValues();
    const headers = data[0];
    let auCol     = headers.indexOf('audio_url');
    let fiCol     = headers.indexOf('audio_file_id');

    if (auCol === -1) {
      const next = headers.length + 1;
      sessSh.getRange(1, next    ).setValue('audio_url');
      sessSh.getRange(1, next + 1).setValue('audio_file_id');
      auCol = next - 1;
      fiCol = next;
    }

    const sidCol = headers.indexOf('session_id');
    for (let i = 1; i < data.length; i++) {
      if (data[i][sidCol] === sessionId) {
        sessSh.getRange(i + 1, auCol + 1).setValue(fileUrl);
        sessSh.getRange(i + 1, fiCol + 1).setValue(fileId);
        break;
      }
    }
  }

  return { success: true, fileId, fileUrl, fileName: fname, path: fpath };
}

// ── Drive helper: get or create a folder by name under a parent ───────────────
function _driveFolder(parent, name) {
  const iter = parent
    ? parent.getFoldersByName(name)
    : DriveApp.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

// ── Make email safe as a folder name ─────────────────────────────────────────
function _safeEmail(email) {
  return (email || 'anonymous').replace(/[^a-zA-Z0-9._\-]/g, '_');
}
