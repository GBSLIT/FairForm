// server.js (final - dynamic column mapping + GlobalBaseContactEmail + direct MailCondition value)
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

app.use(express.static('./public'));
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'form.html')));
app.get('/health', (req, res) => res.send('ok'));

const {
  TENANT_ID, CLIENT_ID, CLIENT_SECRET,
  DRIVE_ID, EXCEL_ITEM_ID, PARENT_FOLDER_ID = 'root', TABLE_NAME = 'BoothResponses', PORT = 8080
} = process.env;

const mask = s => s ? `${s.slice(0, 4)}...${s.slice(-4)}` : 'MISSING';

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

const msal = new ConfidentialClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, clientSecret: CLIENT_SECRET }
});

async function getToken() {
  const result = await msal.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
  if (!result?.accessToken) throw new Error('No access token returned');
  return result.accessToken;
}

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
async function graphGet(url, token) {
  return axios.get(`https://graph.microsoft.com/v1.0${url}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
}

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

function normalizeName(s = '') {
  return s.toLowerCase().replace(/\./g, '').replace(/^\s*mr\s+/,'').trim();
}
function getGbEmail(name = '') {
  const n = normalizeName(name);
  const m = {
    'amjad abbas':'amjad@globalbasesourcing.com','azhar abbas':'azhar@globalbasesourcing.com',
    'ted':'ted@globalbasesourcing.com','clark':'purchase5@globalbasesourcing.com',
    'oscar':'purchase1@globalbasesourcing.com','jack':'purchase2@globalbasesourcing.com',
    'zhong':'purchase4@globalbasesourcing.com'
  };
  if (m[n]) return m[n];
  const f = n.split(/\s+/)[0];
  return m[f] || '';
}

async function appendRow(token, record) {
  const colsUrl = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/columns`;
  const colsResp = await graphGet(colsUrl, token);
  const columns = colsResp.data?.value || [];
  const colNames = columns.map(c => c.name);
  if (!colNames.length) throw new Error('Could not read table columns');

  const headerToRecordKey = {
    ID:'ID', FairName:'FairName', CompanyName:'CompanyName', ContactPerson:'ContactPerson',
    ContactEmail:'ContactEmail', MobileNumber:'MobileNumber', Designation:'Designation',
    KeyProductCategory:'KeyProductCategory', CompanyType:'CompanyType',
    YearEstablished:'YearEstablished',              // 👈 ADDED
    Materials:'Materials', FullAddress:'FullAddress', CompanyLocation:'CompanyLocation',
    City:'City', Country:'Country', ProvinceState:'ProvinceState',
    NearestAirport:'NearestAirport', NearestTrain:'NearestTrain',
    GlobalBaseContact:'GlobalBaseContact', GlobalBaseContactEmail:'GlobalBaseContactEmail',
    VisitingCardCount:'VisitingCardCount', BoothPhotoCount:'BoothPhotoCount',
    CatalogueCount:'CatalogueCount', Message:'Message', FolderLink:'FolderLink',
    Timestamp:'Timestamp', 'Status Mail':'StatusMail', MailCondition:'MailCondition'
  };

  const rowValues = colNames.map(h => {
    const key = headerToRecordKey[h];
    const val = key ? record[key] : '';
    if (val === undefined || val === null) return '';
    return Array.isArray(val) ? val.join(', ') : String(val);
  });

  const addUrl = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows/add`;
  await graphPost(addUrl, token, { values: [rowValues] });
}

const fields = upload.fields([
  { name:'visitingCard',maxCount:10 },{ name:'boothPhotos',maxCount:30 },{ name:'catalogue',maxCount:20 }
]);

app.post('/api/submit', fields, async (req, res) => {
  try {
    const tk = await getToken();
    const readableId = 'FORM_' + new Date().toISOString().replace(/[-:.TZ]/g,'').slice(0,14) + '_' + uuidv4().slice(0,6).toUpperCase();
    const folder = await createResponseFolder(tk, readableId, req.body.companyName);

    const vcCount = await uploadGroup(tk, folder.id, req.files['visitingCard']||[]);
    const bpCount = await uploadGroup(tk, folder.id, req.files['boothPhotos']||[]);
    const ctCount = await uploadGroup(tk, folder.id, req.files['catalogue']||[]);

    const statusMail = '';
    const mailCondition = readableId ? (statusMail ? 'No':'Yes') : '';

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
      YearEstablished: (req.body.yearEstablished || '').toString().trim(),   // 👈 ADDED
      Materials: req.body.materials || '',
      FullAddress: req.body.fullAddress || '',
      CompanyLocation: req.body.companyLocation || '',
      City: req.body.city || '',
      Country: req.body.country || '',
      ProvinceState: req.body.provinceState || '',
      NearestAirport: req.body.nearestAirport || '',
      NearestTrain: req.body.nearestTrain || '',
      GlobalBaseContact: req.body.gbContact || '',
      GlobalBaseContactEmail: getGbEmail(req.body.gbContact || ''),
      VisitingCardCount: vcCount,
      BoothPhotoCount: bpCount,
      CatalogueCount: ctCount,
      Message: req.body.message || '',
      FolderLink: folder.webUrl,
      Timestamp: new Date().toISOString(),
      StatusMail: statusMail,
      MailCondition: mailCondition
    };

    await appendRow(tk, record);

    res.json({ ok:true, id:readableId, folder:folder.webUrl,
      counts:{ visitingCard:vcCount, boothPhotos:bpCount, catalogue:ctCount } });

  } catch(e) {
    res.status(500).json({ ok:false, error:e?.response?.data || e?.message || 'unknown server error' });
  }
});

const LISTEN_PORT = process.env.PORT || 8080;
app.listen(LISTEN_PORT, () => console.log(`Server running on ${LISTEN_PORT}`));
