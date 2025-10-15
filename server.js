// server.js (debug-friendly)
import express from 'express';
import multer from 'multer';
import axios from 'axios';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { v4 as uuidv4 } from 'uuid';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config();
const app = express();
const upload = multer({ limits: { fileSize: 100 * 1024 * 1024 } }); // 100 MB

// simple static + root serve
app.use(express.static('./public'));
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'form.html')));
app.get('/health', (req, res) => res.send('ok'));

// Required env values
const {
  TENANT_ID, CLIENT_ID, CLIENT_SECRET,
  DRIVE_ID, EXCEL_ITEM_ID, PARENT_FOLDER_ID = 'root', TABLE_NAME = 'BoothResponses', PORT = 8080
} = process.env;

// mask helper
const mask = s => s ? `${s.slice(0,4)}...${s.slice(-4)}` : 'MISSING';

// Log env (masked) at startup so we can confirm values exist in Render
console.log('--- STARTUP ENV CHECK ---');
console.log('TENANT_ID', mask(TENANT_ID));
console.log('CLIENT_ID', mask(CLIENT_ID));
console.log('CLIENT_SECRET', CLIENT_SECRET ? 'SET' : 'MISSING');
console.log('DRIVE_ID', mask(DRIVE_ID));
console.log('EXCEL_ITEM_ID', mask(EXCEL_ITEM_ID));
console.log('PARENT_FOLDER_ID', mask(PARENT_FOLDER_ID));
console.log('TABLE_NAME', TABLE_NAME);
console.log('PORT', PORT);
console.log('-------------------------');

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !DRIVE_ID || !EXCEL_ITEM_ID) {
  console.error('One or more critical environment variables are missing. Please set TENANT_ID, CLIENT_ID, CLIENT_SECRET, DRIVE_ID, EXCEL_ITEM_ID in Render env vars.');
}

// MSAL
const msal = new ConfidentialClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, clientSecret: CLIENT_SECRET }
});

async function getToken() {
  try {
    const result = await msal.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
    if (!result || !result.accessToken) throw new Error('No access token returned');
    return result.accessToken;
  } catch (err) {
    console.error('Token acquisition error', err?.message || err);
    throw err;
  }
}

async function graphPost(url, token, data) {
  return axios.post(`https://graph.microsoft.com/v1.0${url}`, data, { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }});
}
async function graphPutBinary(url, token, buffer, contentType) {
  return axios.put(`https://graph.microsoft.com/v1.0${url}`, buffer, { headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType || 'application/octet-stream' }});
}

// create folder
async function createResponseFolder(token, readableId, fairName = '') {
  try {
    const safeFair = (fairName || '').trim().replace(/[^\w. ]/g, '_').slice(0, 60);
    const folderName = `${readableId}${safeFair ? '_' + safeFair : ''}`;
    const base = PARENT_FOLDER_ID === 'root'
      ? `/drives/${DRIVE_ID}/root/children`
      : `/drives/${DRIVE_ID}/items/${PARENT_FOLDER_ID}/children`;
    const body = { name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' };
    const resp = await graphPost(base, token, body);
    return resp.data;
  } catch (err) {
    console.error('Create folder error', err?.response?.data || err.message || err);
    throw err;
  }
}

async function uploadGroup(token, folderId, files) {
  let count = 0;
  for (const f of files) {
    try {
      const safeName = f.originalname.replace(/[^\w.\- ]/g, '_');
      const pathUrl = `/drives/${DRIVE_ID}/items/${folderId}:/${encodeURIComponent(safeName)}:/content`;
      await graphPutBinary(pathUrl, token, f.buffer, f.mimetype);
      count++;
    } catch (err) {
      console.error('Upload file error', f.originalname, err?.response?.data || err.message || err);
      throw err;
    }
  }
  return count;
}

async function appendRow(token, row) {
  try {
    const url = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows/add`;
    const values = [[
      row.ID, row.FairName, row.CompanyName, row.ContactPerson, row.ContactEmail,
      row.MobileNumber, row.Designation, row.KeyProductCategory, row.CompanyType,
      row.Materials, row.FullAddress, row.CompanyLocation, row.GlobalBaseContact,
      row.VisitingCardCount, row.BoothPhotoCount, row.CatalogueCount, row.FolderLink, row.Timestamp
    ]];
    await graphPost(url, token, { values });
  } catch (err) {
    console.error('Append row error', err?.response?.data || err.message || err);
    throw err;
  }
}

const fields = upload.fields([
  { name: 'visitingCard', maxCount: 10 },
  { name: 'boothPhotos', maxCount: 30 },
  { name: 'catalogue', maxCount: 20 }
]);

app.post('/api/submit', fields, async (req, res) => {
  try {
    // quick body log
    console.log('Received submit', { body: Object.keys(req.body), fileGroups: Object.keys(req.files || {}) });

    const tk = await getToken();

    const readableId = 'FORM_' + new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0,14) + '_' + uuidv4().slice(0,6).toUpperCase();

    const folder = await createResponseFolder(tk, readableId, req.body.fairName);
    console.log('Folder created', { id: folder?.id, webUrl: folder?.webUrl });

    const vcCount = await uploadGroup(tk, folder.id, req.files['visitingCard'] || []);
    const bpCount = await uploadGroup(tk, folder.id, req.files['boothPhotos'] || []);
    const ctCount = await uploadGroup(tk, folder.id, req.files['catalogue'] || []);

    const record = {
      ID: readableId,
      FairName: req.body.fairName || '',
      CompanyName: req.body.companyName || '',
      ContactPerson: req.body.contactPerson || '',
      ContactEmail: req.body.contactEmail || '',
      MobileNumber: req.body.mobileNumber || '',
      Designation: req.body.designation || '',
      KeyProductCategory: req.body.keyProductCategory || '',
      CompanyType: req.body.companyType || '',
      Materials: req.body.materials || '',
      FullAddress: req.body.fullAddress || '',
      CompanyLocation: req.body.companyLocation || '',
      GlobalBaseContact: req.body.gbContact || '',
      VisitingCardCount: vcCount,
      BoothPhotoCount: bpCount,
      CatalogueCount: ctCount,
      FolderLink: folder.webUrl,
      Timestamp: new Date().toISOString()
    };

    await appendRow(tk, record);

    res.json({ ok: true, id: readableId, folder: folder.webUrl, counts: { visitingCard: vcCount, boothPhotos: bpCount, catalogue: ctCount } });
  } catch (e) {
    // show as much safe detail as possible
    console.error('Submit error full', e?.response?.data || e?.message || e);
    const errBody = e?.response?.data || { message: e?.message || 'unknown server error' };
    // return the Graph error (if any) to the client so you can see it in browser devtools
    res.status(500).json({ ok: false, error: errBody });
  }
});

const LISTEN_PORT = process.env.PORT || 8080;
app.listen(LISTEN_PORT, () => console.log(`Server running on ${LISTEN_PORT}`));
