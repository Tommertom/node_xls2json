const xlsx = require('xlsx');
const fs = require('fs');

function excelToJson(inputFile, outputFile) {
    // Read Excel file
    const workbook = xlsx.readFile(inputFile);

    // Assuming the first sheet is the one you want to convert
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON
    const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });

    // Extract field names from the first row of the data, and trim spaces
    const fieldNames = jsonData[0].map(field => String(field).trim());

    // Remove the first row from the data
    const dataWithoutHeader = jsonData.slice(1);

    // Create an array of objects with field names and values
    const jsonArray = dataWithoutHeader.map(row => {
        const obj = {};

        // initialise obj with empty strings
        fieldNames.forEach(field => {
            obj[field] = '';
        });

        // populate obj with values from the current row
        row.forEach((value, index) => {
            obj[fieldNames[index]] = String(value).trim(); // Convert values to strings
        });
        return obj;
    });

    // Write the JSON data to a file
    fs.writeFileSync(outputFile, JSON.stringify(jsonArray, null, 2));

    console.log('Conversion complete. JSON file created:', outputFile, '\n', 'With number of rows:', jsonArray.length);
}

// Extract command line parameters
const [, , inputFile, outputFile] = process.argv;

// Check if required parameters are provided
if (!inputFile || !outputFile) {
    console.error('Usage: node excelToJson.js <inputFile> <outputFile> [sheetName]');
    process.exit(1);
}

// Example usage
excelToJson(inputFile, outputFile);
