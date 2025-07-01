import docxTables from 'docx-tables';

docxTables({ file: 'c:/temp/simple.docx' })
  .then((data) => {
    // 'data' will contain the extracted table data in JSON format
    console.log(data);
  })
  .catch((error) => {
    console.error(error);
  });