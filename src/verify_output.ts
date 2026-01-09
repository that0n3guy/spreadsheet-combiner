import * as xlsx from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, '../output');
const files = fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.startsWith('combined_'))
    .map(f => ({ name: f, time: fs.statSync(path.join(OUTPUT_DIR, f)).mtime.getTime() }))
    .sort((a, b) => b.time - a.time);

const combinedFile = files[0]?.name;

if (!combinedFile) {
    console.error('No combined file found');
    process.exit(1);
}

const filePath = path.join(OUTPUT_DIR, combinedFile);
console.log(`Verifying file: ${filePath}`);
const wb = xlsx.readFile(filePath);
console.log('Sheet Names:', wb.SheetNames);
