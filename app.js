const express = require('express');
const app = module.exports = express();
const path = require('path');
const docx = require("docx")

const {
    Document,
    Packer,
    Paragraph,
    TableCell,
    TableRow
} = docx;

app.use(express.static('static', {
    maxage: false,
    etag: false
}))

app.set('storage', path.join(__dirname, 'storage'));

// app.get('/', (req, res) => {

//     const doc = new Document();

//     const table = doc.createTable(3, 3)
//     table.set
//     // table.addChildElement(new TableRow())
//     // table.addChildElement(new TableRow())
//     doc.addTable(table)

//     const packer = new Packer();

//     packer.toBase64String(doc).then((b64string) => {
//         res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
//         res.send(Buffer.from(b64string, 'base64'));
//     })

// })

app.listen(3000, () => {
    console.log('Start app')
});