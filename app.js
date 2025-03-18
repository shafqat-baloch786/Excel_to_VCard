const ExcelJS = require('exceljs');
const vCard = require('vcards-js');
const fs = require('fs');
const path = require('path');

// Excel file
const workbook = new ExcelJS.Workbook();
<<<<<<< HEAD
const excelFilePath = './Excel_file.xlsx';
const outputFileName = 'vcards.vcf';

=======
const excelFilePath = './BSCS SP - 19.xlsx';
const outputFileName = 'vcards.vcf';
>>>>>>> 95ce756 (First commit)
workbook.xlsx.readFile(excelFilePath)
    .then(() => {
        // Accessing file
        const worksheet = workbook.getWorksheet(1);

        // Initializing the array to store vCard data
        const vCards = [];

        // Iterating through all rows
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            // Fetching data from row A (contact name)
            const contactName = row.getCell('A').value;

            // Fetching data from row E (contact number)
            const contactNumber = row.getCell('E').value;

            // Creating vCard
            if (contactName && contactNumber) {
                // Create VCard
                const vCardData = vCard();
                vCardData.firstName = contactName;
                vCardData.cellPhone = contactNumber;

                // Adding VCard data to the array
                vCards.push(vCardData.getFormattedString());
            }
        });

        // Saving VCards to the file in same folder
        const outputFilePath = path.join(__dirname, outputFileName);
        fs.writeFileSync(outputFilePath, vCards.join('\n\n'));
        console.log(`VCard file created successfully at: ${outputFilePath}`);
    })
    .catch(error => {
        console.error('Error reading Excel file:', error.message);
    });
