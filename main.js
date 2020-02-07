const xlsx = require('node-xlsx');
const fs = require('fs');

const targetSheets = ['YOC填报(直营承包)'];
const targetItems = ['转让费（工程服务类）', '承包方支付方式：']; // will output sheet by sheet
const shouldOutputPath = true;

const result = [];

async function ls(path) {
  const dir = await fs.promises.opendir(path)
  for await (const dirent of dir) {
    console.log(dirent.name);
    const workbook = xlsx.parse(`resources/${dirent.name}`);
    workbook.forEach(sheet => {
      if (targetSheets.includes(sheet.name) && sheet.data) {
        console.log('sheet: ', sheet.name);
        console.log('sheet row length', sheet.data.length);
        // find correct sheet, iterate h=sheet.data (get arrays), read first item, find the one ea
        sheet.data.forEach(row => {
          if (row.find(cell => targetItems.includes(typeof cell !== 'string' ? cell : cell.trim()))) {
            result.push({
              dir: `${dirent.name} >>> ${sheet.name}`,
              data: row,
            });
          }
        });
      }
    });
  }
  // after iterated all files
  console.log('result: ', result);
  const buffer = xlsx.build([{
    name: 'sheet 1',
    data: result.reduce((acc, cur) => {
      if (shouldOutputPath) acc.push([cur.dir]);
      acc.push(cur.data);
      return acc;
    }, []),
  }]);
  fs.writeFile('dest/test.xlsx', buffer, err => {
    if (err) {
      return console.log(err);
    }
    console.log('The file is saved!');
  });
}

ls('resources').catch(console.error)
