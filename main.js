#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const csv = require('csvtojson');
const SheetNamingScheme = require('./SheetNamingScheme.json');

const args = process.argv.slice(2);

// Argument deinitions:
//
// arg[0]: Folder with all the csv files.
// arg[1]: Output file name.

args[0] = args[0] || "";
args[1] = args[1] || "Output.xlsx";

if(!args[0]) {
    console.log('Please provide a directory path.');
    process.exit();
}

const inputFolderPath = args[0];
const outputFileName = args[1];

const filesPath = path.resolve(inputFolderPath);
const filesPaths = fs.readdirSync(filesPath).filter(file => {
    return file.match('.csv');
}).reverse();

console.log(filesPaths.length, 'csv files.');

async function main() {

    const workbook = XLSX.utils.book_new();

    const offset = SheetNamingScheme.offset;
    const categories = SheetNamingScheme.categories;

    let index = 0;
    for(let fileName of filesPaths) {
        const filePath = path.resolve(filesPath, fileName);
        const csvString = fs.readFileSync(filePath).toString();
        const csvJsonData = await csv().fromString(csvString);
        const sheet = XLSX.utils.json_to_sheet(csvJsonData);

        sheetName = fileName.replace('.csv', '');
        if(!SheetNamingScheme.useFileNames) {
            sheetName = `${(offset + Math.floor(index / 12))}_${categories[index % categories.length]}`;
        }

        XLSX.utils.book_append_sheet(workbook, sheet, sheetName);

        index++;
    }
    
    console.log('Sheets:', Object.keys(workbook.Sheets));

    const write_opts = {
        type: 'string',
        bookType: 'xlsx',
    }

    XLSX.writeFile(workbook, outputFileName, write_opts);
}

main();
