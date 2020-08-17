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

async function getFilesInDir(dirPath) {
    return new Promise((resolve) => {
        fs.readdir(dirPath, null, (err, filesPaths) => {
            filesPaths = filesPaths.sort((a, b) => {
                return fs.statSync(dirPath + "/" + a).mtime.getTime() - 
                       fs.statSync(dirPath + "/" + b).mtime.getTime();
            });

            console.log('Found', filesPaths.length, '.csv files:');
            
            for(let file of filesPaths) {
                console.log(file);
            }

            resolve(filesPaths);
        });
    })
}

async function convertFilesToWorkbook(dirPath, filesArray) {
    const workbook = XLSX.utils.book_new();

    const sheetNames = [];

    const offset = SheetNamingScheme.offset;
    const categories = SheetNamingScheme.categories;

    let fileIndex = 0;
    for(let fileName of filesArray) {
        const filePath = path.resolve(dirPath, fileName);
        const csvString = fs.readFileSync(filePath).toString();
        const csvJsonData = await csv().fromString(csvString);
        const sheet = XLSX.utils.json_to_sheet(csvJsonData);

        const mkaeSheetName = index => {
            let sheetName = fileName.replace('.csv', '').substring(0, 31);

            if(!SheetNamingScheme.useFileNames) {
                sheetName = `${(offset + Math.floor(fileIndex / 12))} ${categories[fileIndex % categories.length]}`;
            }

            if(index) {
                sheetName = sheetName.substring(0, 28) + ' ' + index;
            }

            return sheetName;
        }

        // check for duplicate sheet names
        let sheetName = mkaeSheetName();
        let sheetIndex = 0;

        if(sheetNames.indexOf(sheetName) != -1) {
            while(sheetNames.indexOf(sheetName) != -1) {
                sheetName = mkaeSheetName(sheetIndex);
                sheetIndex++;
            }
        }

        sheetNames.push(sheetName);

        // add sheet to book
        XLSX.utils.book_append_sheet(workbook, sheet, sheetName);

        fileIndex++;
    }
    
    console.log('Sheets:\n');
    for(let sheet of Object.keys(workbook.Sheets)) {
        console.log('  ' + sheet);
    }

    return workbook;
}

async function main() {
    const dirPath = path.resolve(inputFolderPath);
    const filesArray = await getFilesInDir(dirPath);

    console.log('\nConverting to sheets...\n');
    
    const workbook = await convertFilesToWorkbook(inputFolderPath, filesArray);
    
    const write_opts = {
        type: 'string',
        bookType: 'xlsx',
    }
    XLSX.writeFile(workbook, outputFileName, write_opts);
}

main();
