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

const SPREADSHEET_ID = '1Ajfyfxf3X0Qc7DGnz_sA7RA4SQD_nJTfVy5hPDN_U8o';
const APP_VERSION    = '1.0';

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
