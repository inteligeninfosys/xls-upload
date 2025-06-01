const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const oracledb = require('oracledb');
const cors = require('cors');
const fs = require('fs');
const excel = require('excel4node');
const path = require('path');

const app = express();
app.use(cors());
const upload = multer({ dest: 'uploads/' });

const dbConfig = {
    user: process.env.DB_USER || 'ecol',
    password: process.env.DB_PASSWORD || 'ecol',
    connectString: process.env.DB_CONNECT_STRING || '127.0.0.1:1564/ORCLCDB'
  };    

app.post('/uploadxls', upload.single('file'), async (req, res) => {
    const filePath = req.file.path;
    const owner = req.headers['username'] || 'anonymous';

    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.SheetNames[0];
        const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet], { defval: null });

        if (jsonData.length === 0) return res.status(400).send('Empty file');

        const columns = Object.keys(jsonData[0]);

        // Add 'owner' and 'notesrc' fields if not in sheet
        if (!columns.includes('owner')) {
            columns.push('owner');
        }

        if (!columns.includes('notesrc')) {
            columns.push('notesrc');
        }
        
        const insertSQL = `insert into notehis (${columns.join(', ')})
                       values (${columns.map(c => `:${c}`).join(', ')})`;

        const results = [];
        const connection = await oracledb.getConnection(dbConfig);

        for (let i = 0; i < jsonData.length; i++) {
            const bindRow = {};
            //columns.forEach(col => bindRow[col] = jsonData[i][col]);
            columns.forEach(col => {
                if (col === 'owner') {
                    bindRow[col] = owner;
                } else if (col === 'notesrc') {
                    bindRow[col] = 'uploaded a note';
                } else {
                    bindRow[col] = jsonData[i][col];
                }
            });

            try {
                await connection.execute(insertSQL, bindRow, { autoCommit: false });
                results.push({ ...jsonData[i], owner, notesrc: 'uploaded a note', Status: 'Success', Message: '' });
            } catch (err) {
                results.push({ ...jsonData[i], owner, notesrc: 'uploaded a note', Status: 'Failed', Message: err.message });
            }
        }

        await connection.commit();
        await connection.close();

        // Create Excel report
        const wb = new excel.Workbook();
        const ws = wb.addWorksheet('Upload Report');

        const headers = [...columns, 'Status', 'Message'];
        headers.forEach((h, i) => ws.cell(1, i + 1).string(h));

        results.forEach((row, rowIndex) => {
            headers.forEach((h, colIndex) => {
                const val = row[h] || '';
                if (typeof val === 'number') {
                    ws.cell(rowIndex + 2, colIndex + 1).number(val);
                } else {
                    ws.cell(rowIndex + 2, colIndex + 1).string(String(val));
                }
            });
        });

        const outputFilePath = path.join(__dirname, 'uploads', `report_${Date.now()}.xlsx`);
        wb.write(outputFilePath, (err) => {
            if (err) return res.status(500).send('Error writing report.');
            res.download(outputFilePath, () => {
                fs.unlinkSync(outputFilePath); // delete after download
                fs.unlinkSync(filePath);       // clean uploaded file
            });
        });

    } catch (err) {
        console.error('Processing error:', err);
        res.status(500).send('Internal server error');
    }
});

app.listen(3000, () => {
    console.log('Server started on port 3000');
});
