const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

// Load .env.local untuk development lokal
const envPath = path.join(__dirname, '..', '.env.local');
if (fs.existsSync(envPath)) {
  const envContent = fs.readFileSync(envPath, 'utf8');
  envContent.split('\n').forEach(line => {
    const match = line.match(/^([^=]+)=(.*)$/);
    if (match && !process.env[match[1]]) {
      let val = match[2];
      if (val.startsWith('"') && val.endsWith('"')) val = val.slice(1, -1);
      process.env[match[1]] = val;
    }
  });
}

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate=600');

  try {
    const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
    if (!privateKey || !process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL) {
      throw new Error('Missing GOOGLE_PRIVATE_KEY or GOOGLE_SERVICE_ACCOUNT_EMAIL env vars');
    }

    // Autentikasi menggunakan service account
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: privateKey,
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });
    const sheetId = process.env.GOOGLE_SHEET_ID;

    // Ambil metadata untuk mendapatkan nama sheet yang benar
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const existingSheets = meta.data.sheets.map(s => s.properties.title);
    const sheetName = existingSheets[0];
    const reviewSheetName = 'ReviewLog';

    // Helper to ensure sheet exists
    const ensureSheet = async (title, headers) => {
      if (!existingSheets.includes(title)) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: sheetId,
          requestBody: {
            requests: [{ addSheet: { properties: { title } } }]
          }
        });
        // Add headers
        await sheets.spreadsheets.values.update({
          spreadsheetId: sheetId,
          range: `'${title}'!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [headers] }
        });
        existingSheets.push(title); // Update local list
      }
    };

    if (req.method === 'POST') {
      const { type, data, imageFile, reviewData, updateData, email, pass } = req.body;

      if (type === 'login') {
        const adminEmail = process.env.ADMIN_EMAIL;
        const adminPass = process.env.ADMIN_PASSWORD;

        if (email === adminEmail && pass === adminPass) {
          return res.status(200).json({
            success: true,
            user: { role: 'admin', name: 'Administrator', email: adminEmail }
          });
        } else {
          return res.status(401).json({ success: false, error: 'Email atau Password salah!' });
        }
      }

      if (type === 'review' && reviewData) {
        // Ensure ReviewLog sheet exists
        await ensureSheet(reviewSheetName, ['TxKey', 'Status', 'Notes', 'Reviewer', 'Timestamp', 'CorrectName', 'CorrectNik']);

        // Append review status to ReviewLog sheet
        const values = [[
          reviewData.txKey,
          reviewData.status,
          reviewData.notes,
          reviewData.reviewer,
          new Date().toISOString(),
          reviewData.correctName || '',
          reviewData.correctNik || ''
        ]];

        await sheets.spreadsheets.values.append({
          spreadsheetId: sheetId,
          range: `'${reviewSheetName}'!A:G`,
          valueInputOption: 'USER_ENTERED',
          insertDataOption: 'OVERWRITE',
          requestBody: { values },
        });

        return res.status(200).json({ success: true, message: 'Status review berhasil disimpan' });
      }

      if (type === 'updateRow' && updateData) {
        // Update specific row in main sheet (Columns C and E)
        // updateData should have { rowNo, name, nik }
        // We need to find the actual row index. Assuming rowNo is in Col A.
        // For simplicity, we assume rowNo corresponds to the index if no sorting happened, 
        // but better to search or use the value.
        // Let's assume rowNo is the 1-based index including header (so rowNo + 1)
        const rowIndex = parseInt(updateData.rowNo) + 1; 
        
        await sheets.spreadsheets.values.update({
          spreadsheetId: sheetId,
          range: `'${sheetName}'!C${rowIndex}:C${rowIndex}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[updateData.name]] },
        });

        if (updateData.nik) {
          await sheets.spreadsheets.values.update({
            spreadsheetId: sheetId,
            range: `'${sheetName}'!E${rowIndex}:E${rowIndex}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[`'${updateData.nik}`]] },
          });
        }

        return res.status(200).json({ success: true, message: 'Data transaksi berhasil diperbarui' });
      }

      // Existing transaction append logic...
      if (!data || !Array.isArray(data)) {
        return res.status(400).json({ success: false, error: 'Data invalid' });
      }

      let imageLink = null;
      if (imageFile && imageFile.base64) {
        try {
          const stream = require('stream');
          const base64Str = imageFile.base64;
          const mimeMatch = base64Str.match(/^data:(image\/\w+);base64,/);
          const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
          const base64Data = base64Str.replace(/^data:image\/\w+;base64,/, "");
          const buffer = Buffer.from(base64Data, 'base64');
          
          const driveRes = await drive.files.create({
            requestBody: { name: imageFile.name || 'Bukti_TF.jpg', mimeType },
            media: { mimeType, body: stream.Readable.from(buffer) }
          });
          const fileId = driveRes.data.id;
          
          await drive.permissions.create({
            fileId,
            requestBody: { role: 'reader', type: 'anyone' }
          });
          
          const linkRes = await drive.files.get({ fileId, fields: 'webViewLink' });
          imageLink = linkRes.data.webViewLink;
        } catch (uploadErr) {
          console.error('Gagal upload gambar ke Drive:', uploadErr);
        }
      }

      // Format data untuk append
      const values = data.map(row => {
        let ket = row.keterangan || 'Tabungan';
        if (imageLink && ket.includes('Penarikan')) {
          ket += ` | Link TF: ${imageLink}`;
        }
        return [
          row.no,
          row.bulanTahun,
          row.karyawan,
          row.nominal,
          row.nik ? `'${row.nik}` : '',
          ket
        ];
      });

      await sheets.spreadsheets.values.append({
        spreadsheetId: sheetId,
        range: `'${sheetName}'!A:F`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'OVERWRITE',
        requestBody: { values },
      });

      return res.status(200).json({ success: true, message: 'Data berhasil ditambahkan' });
    }

    // GET untuk ambil data
    const queryType = req.query.type;

    if (queryType === 'review') {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: sheetId,
          range: `'${reviewSheetName}'!A:G`,
        });
        const rows = response.data.values || [];
        const reviews = rows.slice(1).map(row => ({
          txKey: row[0],
          status: row[1],
          notes: row[2],
          reviewer: row[3],
          timestamp: row[4],
          correctName: row[5] || '',
          correctNik: row[6] || ''
        }));
        return res.status(200).json({ success: true, data: reviews });
      } catch (err) {
        // Jika sheet ReviewLog belum ada, return array kosong
        return res.status(200).json({ success: true, data: [], note: 'ReviewLog sheet not found' });
      }
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `'${sheetName}'!A:F`,
    });

    const rows = response.data.values || [];
    
    // Baris pertama = header, sisanya data
    const header = rows[0] || [];
    const data = rows.slice(1).map((row, index) => {
      // Perbaiki parsing nominal: hilangkan titik (ribuan), ubah koma jadi titik (desimal)
      let rawNominal = (row[3] || '0').toString();
      rawNominal = rawNominal.replace(/\./g, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
      const nominalVal = parseFloat(rawNominal) || 0;

      return {
        no: row[0] || (index + 1),
        bulanTahun: row[1] || '',
        karyawan: (row[2] || '').trim(),
        nominal: Math.abs(nominalVal),
        nik: (row[4] || '').trim(),
        keterangan: (row[5] || 'Tabungan').trim(),
        jenisPotongan: 'Investasi'
      };
    });

    // Filter data kosong
    const filtered = data.filter(d => d.karyawan && d.nominal > 0);

    res.status(200).json({
      success: true,
      count: filtered.length,
      header,
      data: filtered,
    });

  } catch (error) {
    console.error('Google Sheets API Error Details:', {
      message: error.message,
      stack: error.stack,
      data: error.response ? error.response.data : 'No response data'
    });
    res.status(500).json({
      success: false,
      error: 'Gagal memproses data Google Sheets',
      detail: error.message,
      apiError: error.response ? error.response.data : null
    });
  }
};
