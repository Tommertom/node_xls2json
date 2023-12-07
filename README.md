# Excel to JSON Converter

This Node.js script converts Excel files into JSON format. It utilizes the `xlsx` library for reading Excel files and `fs` for file system operations. The purpose of this script is to facilitate the conversion of Excel data into a more accessible JSON format.

(text generated by ChatGPT)

## Usage

### Prerequisites

Make sure you have Node.js installed on your system. If not, you can download it [here](https://nodejs.org/).

### Installation

You need xlsx package.

```bash
npm install xlsx
```

### Command

```bash
node excelToJson.js <inputFile> <outputFile>
```

_inputFile_: The path to the Excel file you want to convert.
_<outputFile_: The path to the JSON file where the converted data will be saved.

### Example

```bash
node excelToJson.js input.xlsx output.json Sheet1
```

## Minimal Excel Sheet Specifications

For successful conversion, the Excel sheet should adhere to the following specifications:

- The first row is considered as the header containing field names.
- Data starts from the second row.
- Each column should have a header (field name) in the first row.
- Data types are not preserved in the conversion but made into strings.
- Ensure that the Excel sheet is well-formed and does not contain merged cells or complex structures.

## Script Overview

- Read the Excel file using the xlsx library.
- Convert the specified sheet to JSON format.
- Extract field names from the first row and trim spaces.
- Create an array of objects with field names and corresponding values.
- Write the JSON data to the specified output file.

## Notes

- Ensure that the required command line parameters are provided.
- If the sheet name is not provided, it defaults to the first sheet.
- The resulting JSON file will be formatted with two-space indentation.

Feel free to use this script to streamline the process of converting Excel data to JSON in your Node.js projects.
