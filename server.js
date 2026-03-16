/**
 * Proof-of-concept reference implementation for a Bluebeam Studio roundtrip workflow.
 * Intended for evaluation and development reference only.
 */

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const TokenManager = require('./tokenManager');

const app = express();
const PORT = process.env.PORT || 3000;

// -----------------------------------------------------------------------------
// API BASE URLs
// -----------------------------------------------------------------------------
const API_V1 = 'https://api.bluebeam.com/publicapi/v1';
const API_V2 = 'https://api.bluebeam.com/publicapi/v2';

const CLIENT_ID = process.env.BB_CLIENT_ID;
const WEBHOOK_CALLBACK_URL =
  process.env.WEBHOOK_CALLBACK_URL ||
  `http://localhost:${PORT}/webhook/studio-events`;

// -----------------------------------------------------------------------------
// OPTIONAL SAMPLE / REFERENCE CONSTANTS
// These are only used by the example endpoints below.
// Populate them in .env if you want to test those routes.
// -----------------------------------------------------------------------------
const MARKUP_SESSION_ID = process.env.MARKUP_SESSION_ID || '';
const MARKUP_FILE_ID = process.env.MARKUP_FILE_ID || '';
const MARKUP_FILE_NAME = process.env.MARKUP_FILE_NAME || 'Sample Drawing.pdf';

const CLOSEOUT_PROJECT_ID = process.env.CLOSEOUT_PROJECT_ID || '';
const CLOSEOUT_SESSION_ID = process.env.CLOSEOUT_SESSION_ID || '';
const CLOSEOUT_PROJECT_FILE_ID = process.env.CLOSEOUT_PROJECT_FILE_ID || '';

// -----------------------------------------------------------------------------
// TEAMCENTER STUB — simulates what middleware might receive from Teamcenter
// -----------------------------------------------------------------------------
const DEMO_ASSETS_PATH = process.env.DEMO_ASSETS_PATH || './demo-assets';

const TC_STUB = {
  crNumber: 'CR-12345',
  crDescription: 'Sample drawing review for coordination update',
  drawings: [
    {
      name: 'Sample-Drawing.pdf',
      path: `${DEMO_ASSETS_PATH}/Sample-Drawing.pdf`
    }
  ],
  reviewers: [
    {
      email: process.env.DEMO_REVIEWER_EMAIL || 'reviewer@example.com',
      hasStudioAccount: false
    }
  ],
  workflowNode: 'Checker-Major',
  isEtoCategory: true,
  isPsClass4: false,
  sessionEndDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
};

// -----------------------------------------------------------------------------
// IN-MEMORY STATE — tracks active PoC sessions
// -----------------------------------------------------------------------------
let pocState = {
  sessionId: null,
  subscriptionId: null,
  fileIds: [],
  status: 'idle', // idle | triggered | creating | uploading | inviting | active | finalizing | snapshotting | complete | error
  log: [],
  createdAt: null,
  webhookEvents: []
};

function logStep(msg, type = 'info') {
  const entry = {
    time: new Date().toISOString(),
    msg,
    type
  };
  pocState.log.push(entry);
  console.log(`[${type.toUpperCase()}] ${msg}`);
  return entry;
}

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const fetch = (...args) =>
  import('node-fetch').then(({ default: fetch }) => fetch(...args));

const tokenManager = new TokenManager();

// -----------------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------------
function resetPocState() {
  pocState = {
    sessionId: null,
    subscriptionId: null,
    fileIds: [],
    status: 'idle',
    log: [],
    createdAt: null,
    webhookEvents: []
  };
}

function ensureConfigured(value, name) {
  if (!value) {
    throw new Error(`Missing required configuration: ${name}`);
  }
}

function fileExists(filePath) {
  try {
    return fs.existsSync(filePath);
  } catch {
    return false;
  }
}

// -----------------------------------------------------------------------------
// HEALTH CHECK
// -----------------------------------------------------------------------------
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    config: {
      hasClientId: Boolean(CLIENT_ID),
      webhookCallbackUrl: WEBHOOK_CALLBACK_URL
    },
    sampleConfig: {
      hasMarkupSessionId: Boolean(MARKUP_SESSION_ID),
      hasMarkupFileId: Boolean(MARKUP_FILE_ID),
      hasCloseoutProjectId: Boolean(CLOSEOUT_PROJECT_ID),
      hasCloseoutSessionId: Boolean(CLOSEOUT_SESSION_ID),
      hasCloseoutProjectFileId: Boolean(CLOSEOUT_PROJECT_FILE_ID)
    }
  });
});

// =============================================================================
// POC ROUTES — TEAMCENTER ↔ BLUEBEAM INTEGRATION DEMO
// =============================================================================

// GET current PoC state (for UI polling)
app.get('/poc/state', (req, res) => {
  res.json({
    ...pocState,
    tcStub: TC_STUB
  });
});

// Reset PoC state
app.post('/poc/reset', (req, res) => {
  resetPocState();
  logStep('PoC state reset', 'info');
  res.json({ success: true });
});

// -----------------------------------------------------------------------------
// STEP 1+2 — Simulate Teamcenter trigger + middleware receives event
// -----------------------------------------------------------------------------
app.post('/poc/trigger', (req, res) => {
  pocState.status = 'triggered';
  pocState.log = [];

  logStep(
    `Workflow event received: node=${TC_STUB.workflowNode}, CR=${TC_STUB.crNumber}`,
    'info'
  );
  logStep(
    `Middleware parsing event — ${TC_STUB.drawings.length} drawing(s), ${TC_STUB.reviewers.length} reviewer(s)`,
    'info'
  );
  logStep(
    `Reviewer resolution: isEtoCategory=${TC_STUB.isEtoCategory}, isPsClass4=${TC_STUB.isPsClass4} → inviting ${TC_STUB.reviewers
      .map((r) => r.email)
      .join(', ')}`,
    'info'
  );

  res.json({ success: true, state: pocState });
});

// -----------------------------------------------------------------------------
// STEP 3 — Create Studio Session
// -----------------------------------------------------------------------------
app.post('/poc/create-session', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    pocState.status = 'creating';
    logStep('Creating Bluebeam Studio Session...', 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const body = {
      Name: `${TC_STUB.crNumber}_${TC_STUB.drawings[0].name.replace('.pdf', '')}_Review`,
      Notification: true,
      Restricted: true,
      SessionEndDate: TC_STUB.sessionEndDate,
      DefaultPermissions: [
        { Type: 'Markup', Allow: 'Allow' },
        { Type: 'SaveCopy', Allow: 'Allow' },
        { Type: 'PrintCopy', Allow: 'Allow' },
        { Type: 'MarkupAlert', Allow: 'Allow' },
        { Type: 'AddDocuments', Allow: 'Deny' }
      ]
    };

    const response = await fetch(`${API_V1}/sessions`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id: CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Session creation failed: ${response.status} - ${err}`);
    }

    const data = await response.json();
    pocState.sessionId = data.Id;
    pocState.createdAt = new Date().toISOString();

    logStep(`Session created successfully: ID=${pocState.sessionId}`, 'success');

    res.json({
      success: true,
      sessionId: pocState.sessionId,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 4 — Register Webhook Subscription
// -----------------------------------------------------------------------------
app.post('/poc/register-webhook', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) {
      throw new Error('No active session — run create-session first');
    }

    logStep(`Registering webhook for session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const response = await fetch(`${API_V2}/subscriptions`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id: CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        sourceType: 'session',
        resourceId: pocState.sessionId,
        callbackURI: WEBHOOK_CALLBACK_URL
      })
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Webhook registration failed: ${response.status} - ${err}`);
    }

    const data = await response.json();
    pocState.subscriptionId = data.subscriptionId;

    logStep(`Webhook registered: subscriptionId=${pocState.subscriptionId}`, 'success');
    logStep(`Callback URL configured: ${WEBHOOK_CALLBACK_URL}`, 'info');

    res.json({
      success: true,
      subscriptionId: pocState.subscriptionId,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 5 — Upload PDF (3-step: metadata → S3 → confirm)
// -----------------------------------------------------------------------------
app.post('/poc/upload-file', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) {
      throw new Error('No active session');
    }

    pocState.status = 'uploading';

    const drawing = TC_STUB.drawings[0];
    logStep(`Uploading ${drawing.name} to session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    // 5a — Create metadata block
    logStep('Step 5a: Creating metadata block...', 'info');

    const metaResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}/files`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id: CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        Name: drawing.name,
        Source: `teamcenter://cr/${TC_STUB.crNumber}/${drawing.name}`
      })
    });

    if (!metaResp.ok) {
      const err = await metaResp.text();
      throw new Error(`Metadata block failed: ${metaResp.status} - ${err}`);
    }

    const metaData = await metaResp.json();
    const fileId = metaData.Id;
    const uploadUrl = metaData.UploadUrl;
    const uploadContentType = metaData.UploadContentType || 'application/pdf';

    logStep(`Metadata block created: fileId=${fileId}`, 'success');
    logStep('Upload URL received — uploading file to storage...', 'info');

    // 5b — Upload to storage
    logStep('Step 5b: Uploading PDF binary...', 'info');

    let pdfBuffer;
    if (fileExists(drawing.path)) {
      pdfBuffer = fs.readFileSync(drawing.path);
    } else {
      pdfBuffer = Buffer.from(
        '%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj 3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj\nxref\n0 4\n0000000000 65535 f\n0000000009 00000 n\n0000000058 00000 n\n0000000115 00000 n\ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n190\n%%EOF'
      );
      logStep('No sample PDF found — using minimal demo PDF', 'info');
    }

    const s3Resp = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'x-amz-server-side-encryption': 'AES256',
        'Content-Type': uploadContentType
      },
      body: pdfBuffer
    });

    if (!s3Resp.ok) {
      throw new Error(`Binary upload failed: ${s3Resp.status}`);
    }

    logStep('Binary upload complete', 'success');

    // 5c — Confirm upload
    logStep('Step 5c: Confirming upload with Bluebeam...', 'info');

    const confirmResp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/confirm-upload`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID
        }
      }
    );

    if (!confirmResp.ok) {
      const err = await confirmResp.text();
      throw new Error(`Confirm upload failed: ${confirmResp.status} - ${err}`);
    }

    pocState.fileIds.push({
      fileId,
      name: drawing.name,
      source: `teamcenter://cr/${TC_STUB.crNumber}/${drawing.name}`
    });

    logStep(`File confirmed in session: ${drawing.name} (fileId=${fileId})`, 'success');

    res.json({
      success: true,
      fileId,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 6 — Invite Reviewers
// -----------------------------------------------------------------------------
app.post('/poc/invite-reviewers', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) {
      throw new Error('No active session');
    }

    pocState.status = 'inviting';
    logStep(`Inviting ${TC_STUB.reviewers.length} reviewer(s)...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const results = [];

    for (const reviewer of TC_STUB.reviewers) {
      const endpoint = reviewer.hasStudioAccount
        ? `${API_V1}/sessions/${pocState.sessionId}/users`
        : `${API_V1}/sessions/${pocState.sessionId}/invite`;

      const methodLabel = reviewer.hasStudioAccount
        ? 'existing Studio user flow'
        : 'email invitation flow';

      logStep(`Using ${methodLabel} for ${reviewer.email}`, 'info');

      const resp = await fetch(endpoint, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          Email: reviewer.email,
          SendEmail: true,
          Message: `Please review change request ${TC_STUB.crNumber}: ${TC_STUB.crDescription}`
        })
      });

      if (!resp.ok) {
        const err = await resp.text();
        logStep(`Failed to invite ${reviewer.email}: ${resp.status} - ${err}`, 'warn');
        results.push({
          email: reviewer.email,
          success: false,
          error: err
        });
      } else {
        logStep(`Invited: ${reviewer.email}`, 'success');
        results.push({
          email: reviewer.email,
          success: true
        });
      }
    }

    pocState.status = 'active';
    logStep('Session is now active — reviewers have been notified', 'success');
    logStep(`Session ID associated to change request ${TC_STUB.crNumber}: ${pocState.sessionId}`, 'info');

    res.json({
      success: true,
      results,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 7 — Finalize Session
// -----------------------------------------------------------------------------
app.post('/poc/finalize', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) {
      throw new Error('No active session');
    }

    pocState.status = 'finalizing';
    logStep(`Finalizing session ${pocState.sessionId}...`, 'info');
    logStep('Updating Session status to Finalizing', 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const resp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id: CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ Status: 'Finalizing' })
    });

    if (!resp.ok) {
      const err = await resp.text();
      throw new Error(`Finalize failed: ${resp.status} - ${err}`);
    }

    logStep('Session finalized — webhook/event flow can continue', 'success');

    res.json({
      success: true,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 8 — Create + Poll Snapshot, Download merged PDF
// -----------------------------------------------------------------------------
app.post('/poc/snapshot', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId || pocState.fileIds.length === 0) {
      throw new Error('No active session or no files uploaded');
    }

    pocState.status = 'snapshotting';

    const { fileId, name } = pocState.fileIds[0];
    logStep(`Requesting snapshot for ${name} (fileId=${fileId})...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const snapResp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/snapshot`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID
        }
      }
    );

    if (!snapResp.ok) {
      const err = await snapResp.text();
      throw new Error(`Snapshot request failed: ${snapResp.status} - ${err}`);
    }

    logStep('Snapshot requested — polling for completion...', 'info');

    const maxAttempts = 20;
    const pollInterval = 5000;
    let attempts = 0;
    let downloadUrl = null;

    while (attempts < maxAttempts) {
      await new Promise((resolve) => setTimeout(resolve, pollInterval));
      attempts++;

      const pollToken = await tokenManager.getValidAccessToken();
      const pollResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/snapshot`,
        {
          headers: {
            Authorization: `Bearer ${pollToken}`,
            client_id: CLIENT_ID
          }
        }
      );

      if (!pollResp.ok) {
        logStep(`Poll attempt ${attempts} failed: ${pollResp.status}`, 'warn');
        continue;
      }

      const pollData = await pollResp.json();
      logStep(`Poll ${attempts}/${maxAttempts}: Status=${pollData.Status}`, 'info');

      if (pollData.Status === 'Complete') {
        downloadUrl = pollData.DownloadUrl;
        logStep('Snapshot complete — download URL received', 'success');
        break;
      }

      if (pollData.Status === 'Error') {
        throw new Error(
          `Snapshot failed on Bluebeam side: ${pollData.Message || 'unknown error'}`
        );
      }
    }

    if (!downloadUrl) {
      throw new Error(`Snapshot did not complete after ${maxAttempts} attempts`);
    }

    logStep('Downloading merged PDF...', 'info');

    const dlResp = await fetch(downloadUrl);
    if (!dlResp.ok) {
      throw new Error(`Download failed: ${dlResp.status}`);
    }

    const pdfBuffer = await dlResp.buffer();

    const publicDir = path.join(__dirname, 'public');
    if (!fs.existsSync(publicDir)) {
      fs.mkdirSync(publicDir, { recursive: true });
    }

    const outputFileName = `${TC_STUB.crNumber}_Reviewed.pdf`;
    const outputPath = path.join(publicDir, outputFileName);

    fs.writeFileSync(outputPath, pdfBuffer);

    pocState.status = 'complete';

    logStep(`Marked-up PDF saved: ${outputFileName} (${pdfBuffer.length} bytes)`, 'success');
    logStep(`Reviewed document ready for return to source system for ${TC_STUB.crNumber}`, 'info');

    res.json({
      success: true,
      downloadPath: `/${outputFileName}`,
      fileSize: pdfBuffer.length,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 9 — Cleanup (delete webhook + session)
// -----------------------------------------------------------------------------
app.post('/poc/cleanup', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) {
      throw new Error('No active session to clean up');
    }

    const accessToken = await tokenManager.getValidAccessToken();

    if (pocState.subscriptionId) {
      logStep(`Deleting webhook subscription ${pocState.subscriptionId}...`, 'info');

      const subResp = await fetch(`${API_V2}/subscriptions/${pocState.subscriptionId}`, {
        method: 'DELETE',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID
        }
      });

      if (subResp.ok) {
        logStep('Webhook subscription deleted', 'success');
      } else {
        logStep(`Subscription delete returned ${subResp.status}`, 'warn');
      }
    }

    logStep(`Deleting session ${pocState.sessionId}...`, 'info');

    const sessResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id: CLIENT_ID
      }
    });

    if (sessResp.ok) {
      logStep('Session deleted', 'success');
    } else {
      logStep(`Session delete returned ${sessResp.status}`, 'warn');
    }

    logStep('Cleanup complete', 'success');

    pocState.sessionId = null;
    pocState.subscriptionId = null;

    res.json({
      success: true,
      state: pocState
    });
  } catch (err) {
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// WEBHOOK LISTENER — receives Bluebeam Studio events
// -----------------------------------------------------------------------------
app.post('/webhook/studio-events', (req, res) => {
  const payload = req.body || {};

  logStep(
    `Webhook received: ResourceType=${payload.ResourceType || 'unknown'}, EventType=${payload.EventType || 'unknown'}`,
    'webhook'
  );

  pocState.webhookEvents.push({
    ...payload,
    receivedAt: new Date().toISOString()
  });

  if (payload.ResourceType === 'Sessions' && payload.EventType === 'Update') {
    logStep('Session update event detected — middleware could trigger next workflow step', 'webhook');
  }

  res.sendStatus(200);
});

// =============================================================================
// OPTIONAL SAMPLE ENDPOINTS
// These are included as additional examples and require env configuration.
// =============================================================================

app.get('/powerbi/markups', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');
    ensureConfigured(MARKUP_SESSION_ID, 'MARKUP_SESSION_ID');
    ensureConfigured(MARKUP_FILE_ID, 'MARKUP_FILE_ID');

    const accessToken = await tokenManager.getValidAccessToken();

    const response = await fetch(
      `${API_V2}/sessions/${MARKUP_SESSION_ID}/files/${MARKUP_FILE_ID}/markups`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID,
          Accept: 'application/json'
        }
      }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to get markups: ${response.status} - ${errText}`);
    }

    const data = await response.json();
    const markups = data.Markups || data || [];

    const flattened = markups.map((m) => ({
      MarkupId: m.Id || m.markupId || null,
      FileName: MARKUP_FILE_NAME,
      FileId: MARKUP_FILE_ID,
      SessionId: MARKUP_SESSION_ID,
      Type: m.Type || m.type || null,
      Subject: m.Subject || m.subject || null,
      Comment: m.Comment || m.comment || null,
      Author: m.Author || m.displayName || null,
      DateCreated: m.DateCreated || m.created || null,
      DateModified: m.DateModified || m.modified || null,
      Page: m.Page || m.pageNumber || null,
      Status: m.Status || m.status || null,
      Color: m.Color || null,
      Layer: m.Layer || null
    }));

    res.json(flattened);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/closeout/files', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');
    ensureConfigured(CLOSEOUT_SESSION_ID, 'CLOSEOUT_SESSION_ID');
    ensureConfigured(CLOSEOUT_PROJECT_FILE_ID, 'CLOSEOUT_PROJECT_FILE_ID');

    const accessToken = await tokenManager.getValidAccessToken();

    const response = await fetch(`${API_V1}/sessions/${CLOSEOUT_SESSION_ID}/files`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id: CLIENT_ID,
        Accept: 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch files: ${response.status} - ${await response.text()}`);
    }

    const json = await response.json();

    const files = (json.Files || []).map((f) => ({
      fileName: f.Name || 'Unknown File',
      sessionFileId: f.Id,
      projectFileId: CLOSEOUT_PROJECT_FILE_ID
    }));

    res.json(files);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/closeout-file', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');
    ensureConfigured(CLOSEOUT_SESSION_ID, 'CLOSEOUT_SESSION_ID');
    ensureConfigured(CLOSEOUT_PROJECT_ID, 'CLOSEOUT_PROJECT_ID');
    ensureConfigured(CLOSEOUT_PROJECT_FILE_ID, 'CLOSEOUT_PROJECT_FILE_ID');

    const { sessionFileId } = req.body;

    if (!sessionFileId) {
      throw new Error('Missing sessionFileId');
    }

    const accessToken = await tokenManager.getValidAccessToken();

    const updateResp = await fetch(
      `${API_V1}/sessions/${CLOSEOUT_SESSION_ID}/files/${sessionFileId}/checkin`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ Comment: 'Sync from Session before removal' })
      }
    );

    if (!updateResp.ok) {
      throw new Error(`Step 1 failed: ${updateResp.status} - ${await updateResp.text()}`);
    }

    const deleteResp = await fetch(
      `${API_V1}/sessions/${CLOSEOUT_SESSION_ID}/files/${sessionFileId}`,
      {
        method: 'DELETE',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID
        }
      }
    );

    if (!deleteResp.ok) {
      throw new Error(`Step 2 failed: ${deleteResp.status} - ${await deleteResp.text()}`);
    }

    const finalResp = await fetch(
      `${API_V1}/projects/${CLOSEOUT_PROJECT_ID}/files/${CLOSEOUT_PROJECT_FILE_ID}/checkin`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id: CLIENT_ID,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ Comment: 'Automated final check-in' })
      }
    );

    if (!finalResp.ok) {
      throw new Error(`Step 3 failed: ${finalResp.status} - ${await finalResp.text()}`);
    }

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// START SERVER
// -----------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`\nBluebeam Integration PoC running at http://localhost:${PORT}`);

  console.log(`\nHEALTH / STATUS:`);
  console.log(`   GET  /health`);
  console.log(`   GET  /poc/state`);
  console.log(`   POST /poc/reset`);

  console.log(`\nPOC ENDPOINTS (Teamcenter ↔ Bluebeam reference flow):`);
  console.log(`   POST /poc/trigger`);
  console.log(`   POST /poc/create-session`);
  console.log(`   POST /poc/register-webhook`);
  console.log(`   POST /poc/upload-file`);
  console.log(`   POST /poc/invite-reviewers`);
  console.log(`   POST /poc/finalize`);
  console.log(`   POST /poc/snapshot`);
  console.log(`   POST /poc/cleanup`);
  console.log(`   POST /webhook/studio-events`);

  console.log(`\nOPTIONAL SAMPLE ENDPOINTS:`);
  console.log(`   GET  /powerbi/markups`);
  console.log(`   GET  /api/closeout/files`);
  console.log(`   POST /api/closeout-file`);

  console.log(`\nStub CR: ${TC_STUB.crNumber} — ${TC_STUB.crDescription}`);
  console.log(`Reviewer: ${TC_STUB.reviewers[0].email}`);
  console.log(`Set DEMO_REVIEWER_EMAIL in .env to override\n`);
});
