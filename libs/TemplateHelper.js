const fs = require('fs');
const path = require('path');
const rimraf = require("rimraf");
const fsExtra = require('fs-extra')

class TemplateHelper {
    static getEmailFromData(data) {

        var email = '';
        if (typeof data == 'object') {
            email = data.text;
        } else {
            email = data
        }

        return email;
    }


    static clearFolder(directory) {
        fs.readdir(directory, (err, files) => {
            if (err) throw err;

            if (files != 'stopka_pliki') {
                for (const file of files) {
                    if (file != 'stopka_pliki') {
                        fs.unlink(path.join(directory, file), err => {
                            if (err) throw err;
                        });
                    }
                }
            }
        });
    }

    static clearFoldersList(directory) {
        fs.readdir(directory, (err, folders) => {
            if (err) throw err;

            for (const fl of folders) {
                rimraf(path.join(directory, fl), err => {
                    if (err) throw err;
                });

            }
        });
    }

    static copyFolderFiles(directory, scandir) {
        fs.readdir(scandir, (err, files) => {
            // console.log(err)
            for (const f of files) {

                fsExtra.copySync(
                    scandir + '/' + f,
                    directory + '/stopka_pliki/' + f)

            }

        })
    }

}


module.exports = TemplateHelper;