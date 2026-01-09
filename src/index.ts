import * as fs from 'fs';
import * as path from 'path';
import * as xlsx from 'xlsx';

const INPUT_DIR = path.join(__dirname, '../input');
const OUTPUT_DIR = path.join(__dirname, '../output');

const main = () => {
    // Ensure directories exist
    if (!fs.existsSync(INPUT_DIR)) {
        console.error(`Input directory not found: ${INPUT_DIR}`);
        return;
    }
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR);
    }

    // Get list of files
    const files = fs.readdirSync(INPUT_DIR).filter(file => {
        const ext = path.extname(file).toLowerCase();
        return ext === '.csv' || ext === '.xlsx';
    });

    if (files.length === 0) {
        console.log('No .csv or .xlsx files found in input directory.');
        return;
    }

    console.log(`Found ${files.length} files to process.`);

    const newWorkbook = xlsx.utils.book_new();

    files.forEach(file => {
        const filePath = path.join(INPUT_DIR, file);
        try {
            console.log(`Processing: ${file}`);
            const fileWorkbook = xlsx.readFile(filePath);
            const ext = path.extname(file).toLowerCase();

            if (ext === '.csv') {
                const firstSheetName = fileWorkbook.SheetNames[0];
                const sheet = fileWorkbook.Sheets[firstSheetName];

                let sheetName = path.parse(file).name;
                if (sheetName.length > 31) {
                    sheetName = sheetName.substring(0, 31);
                }

                let finalSheetName = sheetName;
                let counter = 1;
                while (newWorkbook.SheetNames.includes(finalSheetName)) {
                    const suffix = `_${counter}`;
                    if (sheetName.length + suffix.length > 31) {
                        finalSheetName = sheetName.substring(0, 31 - suffix.length) + suffix;
                    } else {
                        finalSheetName = sheetName + suffix;
                    }
                    counter++;
                }

                xlsx.utils.book_append_sheet(newWorkbook, sheet, finalSheetName);

            } else {
                // For XLSX, iterate over all sheets
                fileWorkbook.SheetNames.forEach(sourceSheetName => {
                    const sheet = fileWorkbook.Sheets[sourceSheetName];

                    let finalSheetName = sourceSheetName;
                    let counter = 1;

                    // Deduplicate sheet names
                    while (newWorkbook.SheetNames.includes(finalSheetName)) {
                        const suffix = `_${counter}`;
                        // Original name + suffix might be too long, but we must be careful 
                        // because we are basing it on the *source* sheet name here, not the filename.
                        // Excel sheet name max length is 31.
                        let baseName = sourceSheetName;
                        if (baseName.length + suffix.length > 31) {
                            baseName = baseName.substring(0, 31 - suffix.length);
                        }
                        finalSheetName = baseName + suffix;
                        counter++;
                    }

                    xlsx.utils.book_append_sheet(newWorkbook, sheet, finalSheetName);
                });
            }

        } catch (error) {
            console.error(`Error processing file ${file}:`, error);
        }
    });

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const outputFilename = `combined_${timestamp}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, outputFilename);

    xlsx.writeFile(newWorkbook, outputPath);
    console.log(`Successfully created: ${outputPath}`);
};

main();
