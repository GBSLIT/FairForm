// server.js (final - aligned with 23-column Excel table)
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

// Serve public files
app.use(express.static('./public'));
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'form.html')));
app.get('/health', (req, res) => res.send('ok'));

// Required environment variables
const {
  TENANT_ID, CLIENT_ID, CLIENT_SECRET,
  DRIVE_ID, EXCEL_ITEM_ID, PARENT_FOLDER_ID = 'root', TABLE_NAME = 'BoothResponses', PORT = 8080
} = process.env;

// Mask helper for logs
const mask = s => s ? `${s.slice(0, 4)}...${s.slice(-4)}` : 'MISSING';

// Log environment summary
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
  console.error('One or more critical environment variables are missing.');
}

// MSAL Auth
const msal = new ConfidentialClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, clientSecret: CLIENT_SECRET }
});

async function getToken() {
  const result = await msal.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
  if (!result?.accessToken) throw new Error('No access token returned');
  return result.accessToken;
}

// Graph helpers
async function graphPost(url, token, data) {
  return axios.post(`https://graph.microsoft.com/v1.0${url}`, data, {
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
  });
}

async function graphPutBinary(url, token, buffer, contentType) {
  return axios.put(`https://graph.microsoft.com/v1.0${url}`, buffer, {
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType || 'application/octet-stream' }
  });
}

// Create a unique folder for each submission
async function createResponseFolder(token, readableId, fairName = '') {
  const safeFair = (fairName || '').trim().replace(/[^\w. ]/g, '_').slice(0, 60);
  const folderName = `${readableId}${safeFair ? '_' + safeFair : ''}`;
  const base = PARENT_FOLDER_ID === 'root'
    ? `/drives/${DRIVE_ID}/root/children`
    : `/drives/${DRIVE_ID}/items/${PARENT_FOLDER_ID}/children`;
  const body = { name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' };
  const resp = await graphPost(base, token, body);
  return resp.data;
}

// Upload all files for each section
async function uploadGroup(token, folderId, files) {
  let count = 0;
  for (const f of files) {
    const safeName = f.originalname.replace(/[^\w.\- ]/g, '_');
    const pathUrl = `/drives/${DRIVE_ID}/items/${folderId}:/${encodeURIComponent(safeName)}:/content`;
    await graphPutBinary(pathUrl, token, f.buffer, f.mimetype);
    count++;
  }
  return count;
}

// Append a row to the Excel table
async function appendRow(token, row) {
  const url = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows/add`;

  // 👇 EXACTLY 23 columns in order (same as Excel)
  const values = [[
    row.ID,                // 1 ID
    row.FairName,          // 2
    row.CompanyName,       // 3
    row.ContactPerson,     // 4
    row.ContactEmail,      // 5
    row.MobileNumber,      // 6
    row.Designation,       // 7
    row.KeyProductCategory,// 8
    row.CompanyType,       // 9
    row.Materials,         // 10
    row.FullAddress,       // 11
    row.CompanyLocation,   // 12
    row.City,              // 13
    row.Country,           // 14
    row.ProvinceState,     // 15
    row.NearestAirport,    // 16
    row.NearestTrain,      // 17
    row.GlobalBaseContact, // 18
    row.VisitingCardCount, // 19
    row.BoothPhotoCount,   // 20
    row.CatalogueCount,    // 21
    row.Message,           // 22
    row.FolderLink,        // 23
    row.Timestamp          // 24 ← optional, if Excel has 24 columns, else remove
  ]];

  await graphPost(url, token, { values });
}

// Multer file fields
const fields = upload.fields([
  { name: 'visitingCard', maxCount: 10 },
  { name: 'boothPhotos', maxCount: 30 },
  { name: 'catalogue', maxCount: 20 }
]);

// Main submission handler
app.post('/api/submit', fields, async (req, res) => {
  try {
    console.log('Received submit', { body: Object.keys(req.body), fileGroups: Object.keys(req.files || {}) });

    const tk = await getToken();
    const readableId = 'FORM_' + new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14) + '_' + uuidv4().slice(0, 6).toUpperCase();

    // Create folder
    const folder = await createResponseFolder(tk, readableId, req.body.fairName);
    console.log('Folder created', { id: folder?.id, webUrl: folder?.webUrl });

    // Upload all file groups
    const vcCount = await uploadGroup(tk, folder.id, req.files['visitingCard'] || []);
    const bpCount = await uploadGroup(tk, folder.id, req.files['boothPhotos'] || []);
    const ctCount = await uploadGroup(tk, folder.id, req.files['catalogue'] || []);

    // Build full record object for Excel
    const record = {
      ID: readableId,
      FairName: req.body.fairName || '',
      CompanyName: req.body.companyName || '',
      ContactPerson: req.body.contactPerson || '',
      ContactEmail: req.body.contactEmail || '',
      MobileNumber: req.body.phoneFull || '',
      Designation: req.body.designation || '',
      KeyProductCategory: req.body.keyProductCategory || '',
      CompanyType: req.body.companyType || '',
      Materials: req.body.materials || '',
      FullAddress: req.body.fullAddress || '',
      CompanyLocation: req.body.companyLocation || '',
      City: req.body.city || '',
      Country: req.body.country || '',
      ProvinceState: req.body.provinceState || '',
      NearestAirport: req.body.nearestAirport || '',
      NearestTrain: req.body.nearestTrain || '',
      GlobalBaseContact: req.body.gbContact || '',
      VisitingCardCount: vcCount,
      BoothPhotoCount: bpCount,
      CatalogueCount: ctCount,
      Message: req.body.message || '',
      FolderLink: folder.webUrl,
      Timestamp: new Date().toISOString()
    };

    await appendRow(tk, record);

    res.json({
      ok: true,
      id: readableId,
      folder: folder.webUrl,
      counts: { visitingCard: vcCount, boothPhotos: bpCount, catalogue: ctCount }
    });
  } catch (e) {
    console.error('Submit error full', e?.response?.data || e?.message || e);
    const errBody = e?.response?.data || { message: e?.message || 'unknown server error' };
    res.status(500).json({ ok: false, error: errBody });
  }
});

// Start server
const LISTEN_PORT = process.env.PORT || 8080;
app.listen(LISTEN_PORT, () => console.log(`Server running on ${LISTEN_PORT}`));
