// server.js (final - dynamic column mapping + MailCondition formula apply + GlobalBaseContactEmail)
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

// ➕ get helper
async function graphGet(url, token) {
  return axios.get(`https://graph.microsoft.com/v1.0${url}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
}

// Create a unique folder for each submission
async function createResponseFolder(token, readableId, fairOrCompany = '') {
  const safeTail = (fairOrCompany || '').trim().replace(/[^\w. ]/g, '_').slice(0, 60);
  const folderName = `${readableId}${safeTail ? '_' + safeTail : ''}`;
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

// === NEW: GlobalBaseContact -> Email resolver ===
function normalizeName(s = '') {
  return s
    .toLowerCase()
    .replace(/\./g, '')        // remove dots
    .replace(/^\s*mr\s+/,'')   // drop leading "mr "
    .trim();
}
function getGbEmail(displayName = '') {
  const n = normalizeName(displayName);
  const map = {
    'amjad abbas': 'amjad@globalbasesourcing.com',
    'azhar abbas': 'azhar@globalbasesourcing.com',
    'ted':         'ted@globalbasesourcing.com',
    'clark':       'purchase5@globalbasesourcing.com',
    'oscar':       'purchase1@globalbasesourcing.com',
    'jack':        'purchase2@globalbasesourcing.com',
    'zhong':       'purchase4@globalbasesourcing.com'
  };
  if (map[n]) return map[n];
  const first = n.split(/\s+/)[0];
  return map[first] || '';
}

// 🔁 Append row with dynamic column mapping (adapts to added right-side columns)
async function appendRow(token, record) {
  // 1) Read table columns (names + order)
  const colsUrl = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/columns`;
  const colsResp = await graphGet(colsUrl, token);
  const columns = colsResp.data?.value || [];
  const colNames = columns.map(c => c.name);

  if (!colNames.length) {
    throw new Error('Could not read table columns; check TABLE_NAME and workbook access.');
  }

  // 2) Map Excel header -> record key (known columns)
  const headerToRecordKey = {
    ID: 'ID',
    FairName: 'FairName',
    CompanyName: 'CompanyName',
    ContactPerson: 'ContactPerson',
    ContactEmail: 'ContactEmail',
    MobileNumber: 'MobileNumber',
    Designation: 'Designation',
    KeyProductCategory: 'KeyProductCategory',
    CompanyType: 'CompanyType',
    Materials: 'Materials',
    FullAddress: 'FullAddress',
    CompanyLocation: 'CompanyLocation',
    City: 'City',
    Country: 'Country',
    ProvinceState: 'ProvinceState',
    NearestAirport: 'NearestAirport',
    NearestTrain: 'NearestTrain',
    GlobalBaseContact: 'GlobalBaseContact',
    GlobalBaseContactEmail: 'GlobalBaseContactEmail',   // 👈 NEW mapping
    VisitingCardCount: 'VisitingCardCount',
    BoothPhotoCount: 'BoothPhotoCount',
    CatalogueCount: 'CatalogueCount',
    Message: 'Message',
    FolderLink: 'FolderLink',
    Timestamp: 'Timestamp'
    // Any extra right-side manual columns not listed here will be set to "".
  };

  // 3) Build row array matching the table's current size & order
  const rowValues = colNames.map(h => {
    const key = headerToRecordKey[h];
    const val = key ? record[key] : '';
    if (val === undefined || val === null) return '';
    return Array.isArray(val) ? val.join(', ') : String(val);
  });

  // 4) Add the row
  const addUrl = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows/add`;
  await graphPost(addUrl, token, { values: [rowValues] });
}

// 🧮 Apply MailCondition formula to the whole column (ensures new row gets it)
async function applyMailConditionFormula(token) {
  const tablePath = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}`;

  // Find the "MailCondition" column
  const cols = await graphGet(`${tablePath}/columns`, token);
  const col = (cols.data?.value || []).find(c => (c.name || '').trim() === 'MailCondition');
  if (!col) {
    console.warn('MailCondition column not found in table; skipping formula apply.');
    return;
  }

  // Get data body range for that column
  const rangeResp = await graphGet(`${tablePath}/columns/${encodeURIComponent(col.id)}/dataBodyRange`, token);
  const address = rangeResp.data?.address;           // e.g. "Sheet1!Y2:Y9999"
  const rowCount = rangeResp.data?.rowCount || 0;    // number of data rows in the table
  if (!address || !rowCount) {
    console.warn('MailCondition dataBodyRange not found or empty; skipping.');
    return;
  }

  // Build formulas matrix to match the column range size (rowCount x 1)
  const formula = '=IF([@ID]="","",IF([@[Status Mail]]="","Yes","No"))';
  const matrix = Array.from({ length: rowCount }, () => [formula]);

  // PATCH formulas to the column range
  const sheetName = address.split('!')[0].replace(/'/g, '');
  const encodedSheet = encodeURIComponent(sheetName);
  await axios.patch(
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/worksheets/${encodedSheet}/range(address='${encodeURIComponent(address)}')`,
    { formulas: matrix },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  );
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

    // Create folder (uses company name as requested earlier)
    const folder = await createResponseFolder(tk, readableId, req.body.companyName);
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
      GlobalBaseContactEmail: getGbEmail(req.body.gbContact || ''),  // 👈 NEW computed email
      VisitingCardCount: vcCount,
      BoothPhotoCount: bpCount,
      CatalogueCount: ctCount,
      Message: req.body.message || '',
      FolderLink: folder.webUrl,
      Timestamp: new Date().toISOString()
    };

    // Add the row
    await appendRow(tk, record);

    // Ensure MailCondition formula is applied to the column (covers the new row too)
    try {
      await applyMailConditionFormula(tk);
    } catch (e) {
      console.warn('applyMailConditionFormula error', e?.response?.data || e.message || e);
    }

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
