const express = require('express');
const fileUpload = require('express-fileupload');
const path = require('path');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

app.use(fileUpload());
app.use(express.static('public'));
app.set('view engine', 'ejs');

const mandatoryFields = [
    'FirstName', 'LastName', 'PolicyID', 'GroupNumber', 'MemberID', 'EntityName', 'BilledAmount',
    'DateOfService', 'TimeOfService', 'SubmissionDate', 'TimeOfSubmission', 'ClaimStatus', 'ClaimType',
    'DeniedReason', 'DateOfAdjudication', 'TimeOfAdjudication', 'DateOfOralNotification', 'TimeOfOralNotification',
    'DescriptionOfService', 'Appealed'
];

const isEmpty = (value) => {
    return value === undefined || value === null || value === '';
};

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/upload', (req, res) => {
    if (!req.files || Object.keys(req.files).length === 0) {
        return res.status(400).send('No files were uploaded.');
    }

    const file = req.files.excelFile;
    const workbook = xlsx.read(file.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = rows[0];
    const dataRows = rows.slice(1);
    const errors = [];
    const results = [];

    dataRows.forEach((row, rowIndex) => {
        const rowErrors = [];
        const rowData = {};

        headers.forEach((header, colIndex) => {
            const value = row[colIndex];
            rowData[header] = value;

            if (mandatoryFields.includes(header) && isEmpty(value)) {
                rowErrors.push(`Row ${rowIndex + 2}, Column ${header}: Value is required.`);
            }
        });

        if (rowErrors.length > 0) {
            errors.push(...rowErrors);
        } else {
            results.push(rowData);
        }
    });

    res.render('result', { results, errors });
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
