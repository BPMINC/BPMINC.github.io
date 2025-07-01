// import docxTables from 'docx-tables';

// docxTables({ file: 'c:/temp/simple.docx' })
//   .then((data) => {
//     // 'data' will contain the extracted table data in JSON format
//     console.log(data);
//   })
//   .catch((error) => {
//     console.error(error);
//   });




// import {fromEvent} from 'file-selector';

// // Open file picker
// const handles = await window.showOpenFilePicker({multiple: true});
// // Get the files
// const files = await fromEvent(handles);
// console.log(files);

import * as fs from 'node:fs';

let filePath = "/home/mysystem/dev/myproject/sayHello.txt";
let newFile = fs.readFileSync(filePath);