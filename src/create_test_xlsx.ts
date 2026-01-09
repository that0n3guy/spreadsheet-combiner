import * as xlsx from 'xlsx';
import * as path from 'path';

const wb = xlsx.utils.book_new();
const data = [
    ['Product', 'Price', 'Stock'],
    ['Widget A', 10, 100],
    ['Widget B', 20, 50]
];
const ws = xlsx.utils.aoa_to_sheet(data);
xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');

const data2 = [
    ['Extra', 'Info'],
    ['More', 'Data']
];
const ws2 = xlsx.utils.aoa_to_sheet(data2);
xlsx.utils.book_append_sheet(wb, ws2, 'Sheet2');


xlsx.writeFile(wb, path.join(__dirname, '../input/test_inventory.xlsx'));
console.log('Created input/test_inventory.xlsx');
