// server.js
// npm i express multer axios msal uuid dotenv
import express from 'express';
import multer from 'multer';
import axios from 'axios';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { v4 as uuidv4 } from 'uuid';
import dotenv from 'dotenv';

dotenv.config();
const app = express();
const upload = multer({ limits: { fileSize: 100 * 1024 * 1024 } }); // up to 100 MB per file

// Environment
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const DRIVE_ID = process.env.DRIVE_ID; // OneDrive or SharePoint drive id
const EXCEL_ITEM_ID = process.env.EXCEL_ITEM_ID; // workbook item id
const PARENT_FOLDER_ID = process.env.PARENT_FOLDER_ID || 'root'; // parent folder where response folders are created
const TABLE_NAME = process.env.TABLE_NAME || 'BoothResponses';

// MSAL confidential client
const msal = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET
  }
});

async function token() {
  const t = await msal.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
  return t.accessToken;
}

async function gPost(url, tk, data, headers = {}) {
  return axios.post(`https://graph.microsoft.com/v1.0${url}`, data, {
    headers: { Authorization: `Bearer ${tk}`, 'Content-Type': 'application/json', ...headers }
  });
}
async function gPutBinary(url, tk, buffer, contentType) {
  return axios.put(`https://graph.microsoft.com/v1.0${url}`, buffer, {
    headers: { Authorization: `Bearer ${tk}`, 'Content-Type': contentType || 'application/octet-stream' }
  });
}

async function createResponseFolder(tk, readableId, fairName) {
  const safeFair = (fairName || '').trim().replace(/[^\w. ]/g, '_').slice(0, 60);
  const folderName = `${readableId}${safeFair ? '_' + safeFair : ''}`;
  const base = PARENT_FOLDER_ID === 'root'
    ? `/drives/${DRIVE_ID}/root/children`
    : `/drives/${DRIVE_ID}/items/${PARENT_FOLDER_ID}/children`;
  const body = { name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' };
  const r = await gPost(base, tk, body);
  return r.data; // driveItem with id and webUrl
}

async function uploadGroup(tk, folderId, files) {
  let count = 0;
  for (const f of files) {
    const safeName = f.originalname.replace(/[^\w.\- ]/g, '_');
    const path = `/drives/${DRIVE_ID}/items/${folderId}:/${encodeURIComponent(safeName)}:/content`;
    await gPutBinary(path, tk, f.buffer, f.mimetype);
    count++;
  }
  return count;
}

async function appendRow(tk, row) {
  const url = `/drives/${DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows/add`;
  await gPost(url, tk, { values: [[
    row.ID,
    row.FairName,
    row.CompanyName,
    row.ContactPerson,
    row.ContactEmail,
    row.MobileNumber,
    row.Designation,
    row.KeyProductCategory,
    row.CompanyType,
    row.Materials,
    row.FullAddress,
    row.CompanyLocation,
    row.GlobalBaseContact,
    row.VisitingCardCount,
    row.BoothPhotoCount,
    row.CatalogueCount,
    row.FolderLink,
    row.Timestamp
  ]]});
}

// accept multiple named file groups to mirror your form
const fields = upload.fields([
  { name: 'visitingCard', maxCount: 10 },
  { name: 'boothPhotos', maxCount: 30 },
  { name: 'catalogue', maxCount: 20 }
]);

app.post('/api/submit', fields, async (req, res) => {
  try {
    const tk = await token();

    const readableId = 'FORM_' + new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0,14) + '_' + uuidv4().slice(0,6).toUpperCase();

    const folder = await createResponseFolder(tk, readableId, req.body.fairName);

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
    console.error(e?.response?.data || e.message);
    res.status(500).json({ ok: false, error: 'Server error' });
  }
});

app.use(express.static('./public')); // place form.html in ./public
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log('Server running on', PORT));
