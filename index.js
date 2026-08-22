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
    let connection = null;
    let filePath = null;

    const uploadType = String(
        req.body.uploadType || ''
    ).trim().toUpperCase();

    const allowPartial =
        String(
            req.body.allowPartial || ''
        ).toLowerCase() === 'true';

    const safeUnlink = (file) => {
        try {
            if (file && fs.existsSync(file)) {
                fs.unlinkSync(file);
            }
        } catch (err) {
            console.error('File cleanup error:', err);
        }
    };

    try {
        if (!req.file) {
            return res.status(400).send('No Excel file was uploaded.');
        }

        filePath = req.file.path;

        const owner = req.headers['username'] || 'anonymous';

        const uploadType =
            String(req.body.uploadType || '')
                .trim()
                .toUpperCase();

        const allowedUploadTypes = [
            'LOAN',
            'CREDIT_CARD'
        ];

        if (!allowedUploadTypes.includes(uploadType)) {

            safeUnlink(filePath);

            return res.status(400).send(
                'Invalid upload type. Select either LOAN or CREDIT_CARD.'
            );
        }

        /*
         * --------------------------------------------------
         * Read Excel
         * --------------------------------------------------
         */
        const workbook = xlsx.readFile(filePath);

        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            safeUnlink(filePath);
            return res.status(400).send('Excel workbook contains no sheets.');
        }

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        let jsonData = xlsx.utils.sheet_to_json(worksheet, {
            defval: null,
            raw: false
        });

        if (!jsonData || jsonData.length === 0) {
            safeUnlink(filePath);
            return res.status(400).send('Excel file contains no data rows.');
        }

        /*
         * --------------------------------------------------
         * Normalize Excel column names
         *
         * Example:
         * " CIF " -> "cif"
         * "ACCNUMBER" -> "accnumber"
         * --------------------------------------------------
         */
        jsonData = jsonData.map(row => {
            const normalizedRow = {};

            Object.keys(row).forEach(key => {
                const normalizedKey = String(key)
                    .trim()
                    .toLowerCase();

                normalizedRow[normalizedKey] = row[key];
            });

            return normalizedRow;
        });

        /*
         * --------------------------------------------------
         * Validate template
         * --------------------------------------------------
         */
        const requiredColumns = [
            'accnumber',
            'cif',
            'notemade'
        ];

        const uploadedColumns = Object.keys(jsonData[0]);

        const missingColumns = requiredColumns.filter(
            column => !uploadedColumns.includes(column)
        );

        if (missingColumns.length > 0) {
            safeUnlink(filePath);

            return res.status(400).send(
                `Invalid Excel template. Missing required column(s): ${missingColumns.join(', ')}. ` +
                `Required columns are: accnumber, cif, notemade.`
            );
        }

        /*
         * Reject unexpected fields.
         *
         * This prevents somebody adding arbitrary columns and
         * expecting them to be inserted into NOTEHIS.
         */
        const unexpectedColumns = uploadedColumns.filter(
            column => !requiredColumns.includes(column)
        );

        if (unexpectedColumns.length > 0) {
            safeUnlink(filePath);

            return res.status(400).send(
                `Invalid Excel template. Unexpected column(s): ${unexpectedColumns.join(', ')}. ` +
                `Only accnumber, cif and notemade are allowed.`
            );
        }

        /*
         * --------------------------------------------------
         * Row limit
         * --------------------------------------------------
         */
        const MAX_ROWS = 50000;

        if (jsonData.length > MAX_ROWS) {
            safeUnlink(filePath);

            return res.status(400).send(
                `Excel file contains ${jsonData.length} rows. Maximum allowed is ${MAX_ROWS}.`
            );
        }

        /*
         * --------------------------------------------------
         * FIXED NOTEHIS INSERT
         *
         * IMPORTANT:
         *
         * Excel cif -> Oracle NOTEHIS.custnumber
         * --------------------------------------------------
         */
        const insertSQL = `
            INSERT INTO notehis
            (
                accnumber,
                custnumber,
                notemade,
                owner,
                notesrc
            )
            VALUES
            (
                :accnumber,
                :custnumber,
                :notemade,
                :owner,
                :notesrc
            )
        `;


        /*
        * --------------------------------------------------
        * Validation rules
        * --------------------------------------------------
        */

        let accnumberRegex;
        let accnumberValidationMessage;

        let cifRegex;
        let cifValidationMessage;

        if (uploadType === 'LOAN') {

            accnumberRegex = /^[0-9]{14}$/;

            accnumberValidationMessage =
                'accnumber must contain exactly 14 digits (0-9) for Loan uploads';

            cifRegex = /^[0-9]{9}$/;

            cifValidationMessage =
                'cif must contain exactly 9 digits (0-9) for Loan uploads';

        } else if (uploadType === 'CREDIT_CARD') {

            accnumberRegex = /^[0-9]{5,20}$/;

            accnumberValidationMessage =
                'accnumber must contain between 5 and 14 digits (0-9) for Credit Card uploads';

            cifRegex = /^[0-9]{5,14}$/;

            cifValidationMessage =
                'cif must contain between 5 and 14 digits (0-9) for Credit Card uploads';
        } else {

            safeUnlink(filePath);

            return res.status(400).send(
                'Invalid upload type. Select either LOAN or CREDIT_CARD.'
            );
        }
        /*
        * --------------------------------------------------
        * PHASE 1:
        * Validate ALL Excel rows before touching Oracle
        * --------------------------------------------------
        */
        const validRows = [];
        const validationFailures = [];

        console.log('[BulkNotes] Starting validation', {
            file: req.file.originalname,
            uploadType,
            allowPartial,
            owner,
            rowsReceived: jsonData.length
        });

        for (let i = 0; i < jsonData.length; i++) {

            const row = jsonData[i];
            const excelRowNumber = i + 2;

            const accnumber =
                row.accnumber !== null &&
                    row.accnumber !== undefined
                    ? String(row.accnumber).trim()
                    : '';

            const cif =
                row.cif !== null &&
                    row.cif !== undefined
                    ? String(row.cif).trim()
                    : '';

            const notemade =
                row.notemade !== null &&
                    row.notemade !== undefined
                    ? String(row.notemade).trim()
                    : '';


            /*
             * Ignore completely blank rows.
             */
            if (!accnumber && !cif && !notemade) {

                console.log(
                    `[BulkNotes] Excel row ${excelRowNumber} ignored - blank row`
                );

                continue;
            }


            const validationErrors = [];


            /*
             * Account validation
             */
            if (!accnumber) {

                validationErrors.push(
                    'accnumber is required'
                );

            } else if (!accnumberRegex.test(accnumber)) {

                validationErrors.push(
                    accnumberValidationMessage
                );
            }


            /*
             * CIF validation
             */
            if (!cif) {

                validationErrors.push(
                    'cif is required'
                );

            } else if (!cifRegex.test(cif)) {

                validationErrors.push(
                    cifValidationMessage
                );
            }


            /*
             * Note validation
             */
            if (!notemade) {

                validationErrors.push(
                    'notemade is required'
                );
            }


            /*
             * Invalid row
             */
            if (validationErrors.length > 0) {

                const failedRow = {
                    Row: excelRowNumber,
                    UploadType: uploadType,
                    accnumber,
                    cif,
                    notemade,
                    owner,
                    notesrc: 'uploaded a note',
                    Status: 'Failed',
                    Message: validationErrors.join('; ')
                };

                validationFailures.push(failedRow);

                console.warn(
                    `[BulkNotes] Validation failed at Excel row ${excelRowNumber}`,
                    {
                        accnumber,
                        cif,
                        errors: validationErrors
                    }
                );

                continue;
            }


            /*
             * Valid row
             */
            validRows.push({
                Row: excelRowNumber,
                UploadType: uploadType,
                accnumber,
                cif,
                notemade
            });
        }

        /*
        * --------------------------------------------------
        * Pre-validation summary
        * --------------------------------------------------
        */

        const preValidationFailedCount =
            validationFailures.length;

        console.log('[BulkNotes] Validation completed', {
            total: validRows.length + validationFailures.length,
            valid: validRows.length,
            failed: preValidationFailedCount,
            allowPartial
        });


        /*
        * Show actual failed rows when there are failures.
        */
        if (preValidationFailedCount > 0) {

            console.warn(
                '[BulkNotes] Rows failing pre-validation:',
                validationFailures.map(row => ({
                    Row: row.Row,
                    accnumber: row.accnumber,
                    cif: row.cif,
                    Message: row.Message
                }))
            );
        }

        /*
        * --------------------------------------------------
        * If validation errors exist and user DID NOT
        * approve partial processing, stop before Oracle.
        * --------------------------------------------------
        */

        if (
            preValidationFailedCount > 0 &&
            !allowPartial
        ) {

            console.warn(
                '[BulkNotes] Upload cancelled because partial upload was not approved',
                {
                    total:
                        validRows.length +
                        validationFailures.length,
                    valid: validRows.length,
                    failed: preValidationFailedCount,
                    owner,
                    uploadType
                }
            );

            safeUnlink(filePath);

            return res.status(422).send(
                `${preValidationFailedCount} row(s) contain validation errors. ` +
                `Partial upload was not approved. No records were inserted.`
            );
        }

        /*
        * --------------------------------------------------
        * PHASE 2:
        * User approved processing.
        *
        * Start results with the validation failures.
        * --------------------------------------------------
        */

        const results = [
            ...validationFailures
        ];

        /*
        * No valid rows
        */
        if (validRows.length === 0) {

            console.warn(
                '[BulkNotes] No valid rows available for database processing.'
            );

        } else {
            /*
            * Only now connect to Oracle.
            */
            connection =
                await oracledb.getConnection(
                    dbConfig
                );

            console.log(
                `[BulkNotes] Oracle connected. Processing ${validRows.length} valid row(s).`
            );


            /*
            * --------------------------------------------------
            * Insert valid rows
            * --------------------------------------------------
            */

            for (const row of validRows) {

                const bindRow = {

                    accnumber:
                        row.accnumber,

                    /*
                     * Excel CIF -> NOTEHIS.CUSTNUMBER
                     */
                    custnumber:
                        row.cif,

                    notemade:
                        row.notemade,

                    owner:
                        owner,

                    notesrc:
                        'uploaded a note'
                };


                try {

                    await connection.execute(
                        insertSQL,
                        bindRow,
                        {
                            autoCommit: false
                        }
                    );


                    results.push({
                        Row: row.Row,
                        UploadType: uploadType,
                        accnumber: row.accnumber,
                        cif: row.cif,
                        notemade: row.notemade,
                        owner,
                        notesrc: 'uploaded a note',
                        Status: 'Success',
                        Message: ''
                    });


                    console.log(
                        `[BulkNotes] Row ${row.Row} inserted successfully`,
                        {
                            accnumber: row.accnumber,
                            cif: row.cif
                        }
                    );


                } catch (err) {

                    /*
                     * Database failure for THIS row.
                     * Other valid rows continue.
                     */
                    results.push({
                        Row: row.Row,
                        UploadType: uploadType,
                        accnumber: row.accnumber,
                        cif: row.cif,
                        notemade: row.notemade,
                        owner,
                        notesrc: 'uploaded a note',
                        Status: 'Failed',
                        Message: err.message
                    });


                    console.error(
                        `[BulkNotes] Oracle insert failed at Excel row ${row.Row}`,
                        {
                            accnumber: row.accnumber,
                            cif: row.cif,
                            code: err.code,
                            errorNum: err.errorNum,
                            message: err.message,
                            offset: err.offset
                        }
                    );
                }
            }

            /*
            * Commit all successful statements.
            */
            await connection.commit();

            console.log(
                '[BulkNotes] Oracle transaction committed.'
            );


            await connection.close();

            connection = null;

            console.log(
                '[BulkNotes] Oracle connection closed.'
            );
        }

        /*
        * Keep report in original Excel row order.
        */
        results.sort(
            (a, b) =>
                a.Row - b.Row
        );


        /*
         * --------------------------------------------------
         * Calculate result counts
         * --------------------------------------------------
         */
        const successCount = results.filter(
            row => row.Status === 'Success'
        ).length;

        const failedCount = results.filter(
            row => row.Status === 'Failed'
        ).length;

        console.log(
            `Bulk notes upload completed. ` +
            `User=${owner}, ` +
            `Total=${results.length}, ` +
            `Success=${successCount}, ` +
            `Failed=${failedCount}`
        );

        /*
         * --------------------------------------------------
         * Generate Excel Upload Report
         * --------------------------------------------------
         */
        const wb = new excel.Workbook();

        const ws = wb.addWorksheet('Upload Report');

        /*
         * IMPORTANT:
         *
         * Report continues to show "cif" because that is
         * what the user uploaded.
         *
         * We don't expose the internal NOTEHIS field name
         * custnumber to the user.
         */
        const reportHeaders = [
            'Row',
            'UploadType',
            'accnumber',
            'cif',
            'notemade',
            'owner',
            'notesrc',
            'Status',
            'Message'
        ];

        reportHeaders.forEach((header, index) => {
            ws.cell(1, index + 1).string(header);
        });

        results.forEach((row, rowIndex) => {

            reportHeaders.forEach((header, colIndex) => {

                const value =
                    row[header] === null ||
                        row[header] === undefined
                        ? ''
                        : row[header];

                if (typeof value === 'number') {

                    ws
                        .cell(rowIndex + 2, colIndex + 1)
                        .number(value);

                } else {

                    ws
                        .cell(rowIndex + 2, colIndex + 1)
                        .string(String(value));
                }
            });
        });

        /*
         * Make report columns easier to read
         */
        ws.column(1).setWidth(10);
        ws.column(2).setWidth(25);
        ws.column(3).setWidth(20);
        ws.column(4).setWidth(60);
        ws.column(5).setWidth(20);
        ws.column(6).setWidth(25);
        ws.column(7).setWidth(15);
        ws.column(8).setWidth(70);

        const outputFilePath = path.join(
            __dirname,
            'uploads',
            `bulk_notes_report_${Date.now()}.xlsx`
        );

        wb.write(outputFilePath, (err) => {

            if (err) {

                console.error(
                    'Error writing bulk upload report:',
                    err
                );

                safeUnlink(filePath);

                return res
                    .status(500)
                    .send('Error generating upload report.');
            }

            /*
             * Optional useful response headers.
             */
            res.setHeader(
                'X-Upload-Total',
                String(results.length)
            );

            res.setHeader(
                'X-Upload-Success',
                String(successCount)
            );

            res.setHeader(
                'X-Upload-Failed',
                String(failedCount)
            );

            res.download(
                outputFilePath,
                'Bulk_Notes_Upload_Report.xlsx',
                downloadErr => {

                    if (downloadErr) {
                        console.error(
                            'Report download error:',
                            downloadErr
                        );
                    }

                    safeUnlink(outputFilePath);
                    safeUnlink(filePath);
                }
            );
        });

    } catch (err) {

        console.error(
            '[BulkNotes] Fatal upload processing error',
            {
                code: err.code,
                errorNum: err.errorNum,
                message: err.message,
                offset: err.offset,
                stack: err.stack
            }
        );

        if (connection) {

            try {

                await connection.rollback();

                console.warn(
                    '[BulkNotes] Oracle transaction rolled back.'
                );

            } catch (rollbackErr) {

                console.error(
                    '[BulkNotes] Oracle rollback failed',
                    {
                        code: rollbackErr.code,
                        message: rollbackErr.message
                    }
                );
            }


            try {

                await connection.close();

            } catch (closeErr) {

                console.error(
                    '[BulkNotes] Oracle connection close failed',
                    {
                        code: closeErr.code,
                        message: closeErr.message
                    }
                );
            }
        }

        safeUnlink(filePath);

        if (!res.headersSent) {

            return res
                .status(500)
                .send(
                    'Bulk notes upload failed due to an internal server error.'
                );
        }
    }
});

app.listen(3000, () => {
    console.log('Server started on port 3000');
});
