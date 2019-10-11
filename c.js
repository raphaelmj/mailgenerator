const program = require('commander');
const curl = new(require('curl-request'))();
var Excel = require('exceljs');
const fs = require('fs')
const fsExtra = require('fs-extra')
const fileSystem = require('file-system')
const TemplateHelper = require('./libs/TemplateHelper')
var slug = require('slug')
const {
    base64encode,
    base64decode
} = require('nodejs-base64');


program
    .command('create-mails-html')
    .action((from, to) => {



        TemplateHelper.clearFolder(__dirname + '/storage/output/html')

        var workbook = new Excel.Workbook();

        workbook.xlsx.readFile(__dirname + '/storage/lista.xlsx')
            .then(function (worksheet) {



                worksheet.eachSheet((worksheet2, sheetId) => {

                    const regex = /\[\[([a-z0-9\/\-.]*)\]\]/gm;

                    worksheet2.eachRow((row, rowNumber) => {


                        if (rowNumber != 1) {

                            var data = JSON.parse(JSON.stringify(row.values));


                            var email = TemplateHelper.getEmailFromData(data[1])
                            var contents = fs.readFileSync(__dirname + '/storage/template_no-font-new-b64.html').toString();

                            if (email != '') {
                                email = `email: <a href="mailto:${email}" style="color:#000">${email}</a>`
                            }
                            contents = contents.replace('{{email}}', email)

                            contents = contents.replace('{{name}}', data[3] + '<br>')

                            if (data[4]) {
                                if (data[4] != '') {
                                    data[4] = data[4] + '<br>'
                                }
                                contents = contents.replace('{{mobile}}', data[4])
                            } else {
                                contents = contents.replace('{{mobile}}', '')
                            }

                            if (data[5]) {
                                if (data[5] != '') {
                                    data[5] = data[5] + '<br>'
                                }
                                contents = contents.replace('{{phone}}', data[5])
                            } else {
                                contents = contents.replace('{{phone}}', '')
                            }

                            if (data[6]) {
                                if (data[6] != '') {
                                    data[6] = data[6] + '<br>'
                                }
                                contents = contents.replace('{{proffesion}}', data[6])
                            } else {
                                contents = contents.replace('{{proffesion}}', '')
                            }

                            let m;

                            while ((m = regex.exec(contents)) !== null) {
                                // This is necessary to avoid infinite loops with zero-width matches
                                if (m.index === regex.lastIndex) {
                                    regex.lastIndex++;
                                }

                                // The result can be accessed through the `m`-variable.
                                m.forEach((match, groupIndex) => {

                                    var imgBase64;
                                    var sign;

                                    if (groupIndex == 0) {
                                        sign = match;
                                    }

                                    if (groupIndex == 1) {
                                        let encoded = base64encode(__dirname + `/static/${match}`);
                                        // console.log(encoded)
                                        imgBase64 = 'data:image/jpeg;base64,' + encoded;
                                    }

                                    contents = contents.replace(sign, imgBase64)

                                    // console.log(`Found match, group ${groupIndex}: ${match}`);
                                });
                            }

                            var slugName = slug(data[3], '-')
                            fs.writeFileSync(__dirname + `/storage/output/html/mail-${sheetId}${rowNumber}-${slugName}.html`, contents);

                        }


                    })

                })


            });

    })


program
    .command('create-mails-html-invers')
    .action((from, to) => {



        TemplateHelper.clearFolder(__dirname + '/storage/output/html')

        var workbook = new Excel.Workbook();

        workbook.xlsx.readFile(__dirname + '/storage/lista.xlsx')
            .then(function (worksheet) {



                worksheet.eachSheet((worksheet2, sheetId) => {

                    const regex = /\[\[([a-z0-9\/\-.]*)\]\]/gm;

                    worksheet2.eachRow((row, rowNumber) => {


                        if (rowNumber != 1) {

                            var data = JSON.parse(JSON.stringify(row.values));


                            var email = TemplateHelper.getEmailFromData(data[1])
                            var contents = fs.readFileSync(__dirname + '/storage/template_invers.html').toString();

                            contents = contents.replace('{{email}}', email)

                            contents = contents.replace('{{name}}', data[3])

                            if (data[4])
                                contents = contents.replace('{{mobile}}', data[4])
                            else
                                contents = contents.replace('{{mobile}}', '')

                            if (data[5])
                                contents = contents.replace('{{phone}}', data[5])
                            else
                                contents = contents.replace('{{phone}}', '')

                            if (data[6])
                                contents = contents.replace('{{proffesion}}', data[6])
                            else
                                contents = contents.replace('{{proffesion}}', '')


                            let m;

                            while ((m = regex.exec(contents)) !== null) {
                                // This is necessary to avoid infinite loops with zero-width matches
                                if (m.index === regex.lastIndex) {
                                    regex.lastIndex++;
                                }

                                // The result can be accessed through the `m`-variable.
                                m.forEach((match, groupIndex) => {

                                    var imgBase64;
                                    var sign;

                                    if (groupIndex == 0) {
                                        sign = match;
                                    }

                                    if (groupIndex == 1) {
                                        let encoded = base64encode(__dirname + `/static/${match}`);
                                        // console.log(encoded)
                                        imgBase64 = 'data:image/jpeg;base64,' + encoded;
                                    }

                                    contents = contents.replace(sign, imgBase64)

                                    // console.log(`Found match, group ${groupIndex}: ${match}`);
                                });
                            }

                            var slugName = slug(data[3], '-')
                            fs.writeFileSync(__dirname + `/storage/output/html_invers/mail-${sheetId}${rowNumber}-${slugName}.html`, contents);

                        }


                    })

                })


            });

    })


program
    .command('create-mails-txt')
    .action((from, to) => {


        TemplateHelper.clearFolder(__dirname + '/storage/output/txt')

        var workbook = new Excel.Workbook();

        workbook.xlsx.readFile(__dirname + '/storage/lista.xlsx')
            .then(function (worksheet) {



                worksheet.eachSheet((worksheet2, sheetId) => {

                    const regex = /\[\[([a-z0-9\/\-.]*)\]\]/gm;

                    worksheet2.eachRow((row, rowNumber) => {


                        if (rowNumber != 1) {

                            var data = JSON.parse(JSON.stringify(row.values));


                            var email = TemplateHelper.getEmailFromData(data[1])
                            var contents = fs.readFileSync(__dirname + '/storage/text.txt').toString();

                            contents = contents.replace('{{email}}', email)

                            contents = contents.replace('{{name}}', data[3])

                            if (data[4])
                                contents = contents.replace('{{mobile}}', data[4])
                            else
                                contents = contents.replace('{{mobile}}', '')

                            if (data[5])
                                contents = contents.replace('{{phone}}', data[5])
                            else
                                contents = contents.replace('{{phone}}', '')

                            if (data[6])
                                contents = contents.replace('{{proffesion}}', data[6])
                            else
                                contents = contents.replace('{{proffesion}}', '')


                            var slugName = slug(data[3], '-')
                            fs.writeFileSync(__dirname + `/storage/output/txt/mail-${sheetId}${rowNumber}-${slugName}.txt`, contents);

                        }


                    })

                })


            });

    })

program
    .command('create-mails-word')
    .action((from, to) => {
        TemplateHelper.clearFolder(__dirname + '/storage/output/html_word')

        var workbook = new Excel.Workbook();

        workbook.xlsx.readFile(__dirname + '/storage/lista.xlsx')
            .then(function (worksheet) {



                worksheet.eachSheet((worksheet2, sheetId) => {

                    const regex = /\[\[([a-z0-9\/\-.]*)\]\]/gm;

                    worksheet2.eachRow((row, rowNumber) => {


                        if (rowNumber != 1) {

                            var data = JSON.parse(JSON.stringify(row.values));


                            var email = TemplateHelper.getEmailFromData(data[1])
                            var contents = fs.readFileSync(__dirname + '/storage/template_word.html').toString();

                            contents = contents.replace('{{email}}', email)

                            contents = contents.replace('{{name}}', data[3])

                            if (data[4])
                                contents = contents.replace('{{mobile}}', data[4])
                            else
                                contents = contents.replace('{{mobile}}', '')

                            if (data[5])
                                contents = contents.replace('{{phone}}', data[5])
                            else
                                contents = contents.replace('{{phone}}', '')

                            if (data[6])
                                contents = contents.replace('{{proffesion}}', data[6])
                            else
                                contents = contents.replace('{{proffesion}}', '')


                            let m;

                            while ((m = regex.exec(contents)) !== null) {
                                // This is necessary to avoid infinite loops with zero-width matches
                                if (m.index === regex.lastIndex) {
                                    regex.lastIndex++;
                                }

                                // The result can be accessed through the `m`-variable.
                                m.forEach((match, groupIndex) => {

                                    var imgBase64;
                                    var sign;

                                    if (groupIndex == 0) {
                                        sign = match;
                                    }

                                    if (groupIndex == 1) {
                                        let encoded = base64encode(__dirname + `/static/${match}`);
                                        // console.log(encoded)
                                        imgBase64 = 'data:image/jpeg;base64,' + encoded;
                                    }

                                    contents = contents.replace(sign, imgBase64)

                                    // console.log(`Found match, group ${groupIndex}: ${match}`);
                                });
                            }

                            var slugName = slug(data[3], '-')
                            fs.writeFileSync(__dirname + `/storage/output/html_word/mail-${sheetId}${rowNumber}-${slugName}.html`, contents);

                        }


                    })

                })


            });
    })

program
    .command('create-mails-word-invers')
    .action((from, to) => {
        TemplateHelper.clearFolder(__dirname + '/storage/output/html_word')

        var workbook = new Excel.Workbook();

        workbook.xlsx.readFile(__dirname + '/storage/lista.xlsx')
            .then(function (worksheet) {



                worksheet.eachSheet((worksheet2, sheetId) => {

                    const regex = /\[\[([a-z0-9\/\-.]*)\]\]/gm;

                    worksheet2.eachRow((row, rowNumber) => {


                        if (rowNumber != 1) {

                            var data = JSON.parse(JSON.stringify(row.values));


                            var email = TemplateHelper.getEmailFromData(data[1])
                            var contents = fs.readFileSync(__dirname + '/storage/template_word_invers.html').toString();

                            contents = contents.replace('{{email}}', email)

                            contents = contents.replace('{{name}}', data[3])

                            if (data[4])
                                contents = contents.replace('{{mobile}}', data[4])
                            else
                                contents = contents.replace('{{mobile}}', '')

                            if (data[5])
                                contents = contents.replace('{{phone}}', data[5])
                            else
                                contents = contents.replace('{{phone}}', '')

                            if (data[6])
                                contents = contents.replace('{{proffesion}}', data[6])
                            else
                                contents = contents.replace('{{proffesion}}', '')


                            let m;

                            while ((m = regex.exec(contents)) !== null) {
                                // This is necessary to avoid infinite loops with zero-width matches
                                if (m.index === regex.lastIndex) {
                                    regex.lastIndex++;
                                }

                                // The result can be accessed through the `m`-variable.
                                m.forEach((match, groupIndex) => {

                                    var imgBase64;
                                    var sign;

                                    if (groupIndex == 0) {
                                        sign = match;
                                    }

                                    if (groupIndex == 1) {
                                        let encoded = base64encode(__dirname + `/static/${match}`);
                                        // console.log(encoded)
                                        imgBase64 = 'data:image/jpeg;base64,' + encoded;
                                    }

                                    contents = contents.replace(sign, imgBase64)

                                    // console.log(`Found match, group ${groupIndex}: ${match}`);
                                });
                            }

                            var slugName = slug(data[3], '-')
                            fs.writeFileSync(__dirname + `/storage/output/html_word_invers/mail-${sheetId}${rowNumber}-${slugName}.html`, contents);

                        }


                    })

                })


            });
    })


program
    .command('create-mails-word-subfolder')
    .action((from, to) => {
        TemplateHelper.clearFoldersList(__dirname + '/storage/output/html_word_subfolder')

        var workbook = new Excel.Workbook();

        workbook.xlsx.readFile(__dirname + '/storage/lista.xlsx')
            .then(function (worksheet) {



                worksheet.eachSheet((worksheet2, sheetId) => {

                    const regex = /\[\[([a-z0-9\/\-.]*)\]\]/gm;

                    worksheet2.eachRow((row, rowNumber) => {


                        if (rowNumber != 1) {

                            var data = JSON.parse(JSON.stringify(row.values));


                            var email = TemplateHelper.getEmailFromData(data[1])
                            var contents = fs.readFileSync(__dirname + '/storage/template_word.html').toString();

                            contents = contents.replace('{{email}}', email)

                            contents = contents.replace('{{name}}', data[3])

                            if (data[4])
                                contents = contents.replace('{{mobile}}', data[4])
                            else
                                contents = contents.replace('{{mobile}}', '')

                            if (data[5])
                                contents = contents.replace('{{phone}}', data[5])
                            else
                                contents = contents.replace('{{phone}}', '')

                            if (data[6])
                                contents = contents.replace('{{proffesion}}', data[6])
                            else
                                contents = contents.replace('{{proffesion}}', '')


                            let m;

                            while ((m = regex.exec(contents)) !== null) {
                                // This is necessary to avoid infinite loops with zero-width matches
                                if (m.index === regex.lastIndex) {
                                    regex.lastIndex++;
                                }

                                // The result can be accessed through the `m`-variable.
                                m.forEach((match, groupIndex) => {

                                    var imgBase64;
                                    var sign;

                                    if (groupIndex == 0) {
                                        sign = match;
                                    }

                                    if (groupIndex == 1) {
                                        let encoded = base64encode(__dirname + `/static/${match}`);
                                        // console.log(encoded)
                                        imgBase64 = 'data:image/jpeg;base64,' + encoded;
                                    }

                                    contents = contents.replace(sign, imgBase64)

                                    // console.log(`Found match, group ${groupIndex}: ${match}`);
                                });
                            }

                            var slugName = slug(data[3], '-')

                            if (!fs.existsSync(__dirname + `/storage/output/html_word_subfolder/${slugName}`)) {
                                fs.mkdirSync(__dirname + `/storage/output/html_word_subfolder/${slugName}`);
                            }
                            TemplateHelper.copyFolderFiles(__dirname + `/storage/output/html_word_subfolder/${slugName}`, __dirname + '/storage/output/stopka_pliki')
                            fs.writeFileSync(__dirname + `/storage/output/html_word_subfolder/${slugName}/stopka.html`, contents);



                        }


                    })

                })


            });
    })

program.parse(process.argv)