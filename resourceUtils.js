const { resolve } = require('path')
const { promisify } = require('util')
const fs = require('fs-extra')
const readdir = promisify(fs.readdir)
const stat = promisify(fs.stat)
var JSZip = require("jszip")
var xlsx = require('xlsx')
var Workbook = require('./workbook').Workbook
const readline = require('readline');
var rowsing = 0

var Utils = {
    askQuestion: function (query){
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout,
        });
    
        return new Promise(resolve => rl.question(query, ans => {
            rl.close();
            resolve(ans);
        }))
    },
    compareProperties: async function(files, destination) {
        try {
            var workbook = new Workbook();
            var fileObj = {
                de: {path:[], name:[]},
                pl:{path:[], name:[]}
            };
            for (let i = 0; i < files.length; i++) {
                var fileName = files[i].split('\\').pop();
                var language = fileName.split('_')[1];
                fileObj[language]['path'].push(files[i]);
                fileObj[language]['name'].push(fileName);
            }
            console.log(fileObj.de.name.length, fileObj.pl.name.length);
            for (let i = 0; i < fileObj.de.name.length; i++) {
                var currentFileName = fileObj.de.name[i];
                global[currentFileName] = workbook.add(currentFileName);

                var defaultFileContent = await new Promise((resolve, reject) => {
                    fs.readFile(fileObj.de.path[i], 'utf8', function(err, data) {
                        rowsArray = data.split('\n');
                        resolve(rowsArray);
                    });
                });
                var compareContent = await new Promise((resolve, reject)=>{
                    fs.readFile(fileObj.pl.path[i], 'utf8', function(err, data) {
                        rowsArray = data.split('\n');
                        resolve(rowsArray);
                    });
                })
                var frKeys = [];
                var compareKeys = [];

                for (let i = 0; i < defaultFileContent.length; i++) {
                    var cleanRow = defaultFileContent[i].replace(/(\r\n|\n|\r)/gm,'');
                    var currentRow = cleanRow.split(/^([^=]+)=/);
                    currentRow.splice(0,1);
                    frKeys.push(currentRow);
                }
                for (let i = 0; i < compareContent.length; i++) {
                    var cleanRowCompare = compareContent[i].replace(/(\r\n|\n|\r)/gm,'');
                    var currentRow = cleanRowCompare.split(/^([^=]+)=/);
                    currentRow.splice(0,1);
                    compareKeys.push(currentRow[0]);
                }

                for (let i = 0; i < frKeys.length; i++) {
                    if (compareKeys.indexOf(frKeys[i][0]) == -1) {
                        var untreatedRow = frKeys[i];

                        if (untreatedRow[0] != '') {
                            global[currentFileName][rowsing][0] = untreatedRow[0];
                            global[currentFileName][rowsing][1] = untreatedRow[1];
                            rowsing++;
                        }
                    }
                }
                rowsing = 0;
            }
            workbook.save(destination);
        } catch (error) {
            console.log(error)
        }
    },
    readAndWriteProperties: async function (files, destination, type) {
        if (type == 'compare') {
            Utils.compareProperties(files, destination);
        } else {
            try {
                var workbook = new Workbook();
                var promiseArr = []
                for (let i = 0; i < files.length; i++) {
                    var sheetName = files[i].split('\\').pop();
                    global[sheetName] = workbook.add(sheetName)
                    var ceva = new Promise((resolve,reject)=>{
                        fs.readFile(files[i], 'utf8', function(err, data) {
                            rowsArray= data.split('\n')
                            resolve(rowsArray);
                        });
                    })
                    promiseArr.push(ceva)
                }
                Promise.all(promiseArr).then((result)=>{
                    for (let r = 0; r < result.length; r++) {
                        var parsedProp = result[r]
                        var sheetName = files[r].split('\\').pop()
                        for (let x = 0; x < parsedProp.length; x++) {
                            let cleanRow = parsedProp[x].replace(/(\r\n|\n|\r)/gm,'')

                            var row = cleanRow.split(/^([^=]+)=/);
                            row.splice(0,1);

                            // alternative solution
                            // let [row, ...rest] = cleanRow.split(/=/);
                            // rest = rest.join('=');
                            // row = [row, rest];

                            // test for arabic and not put them in excel
                            //var arabic = /[\u0600-\u06FF]/;
                            //if (!/\.ar/.test(row[0]) && !arabic.test(row[1]) && row[0] != '') {
                            if (row[0] != '') {
                                global[sheetName][rowsing][0] = row[0]
                                global[sheetName][rowsing][1] = row[1]
                                rowsing++;
                            }
                        }
                        rowsing = 0;
                    }
                    workbook.save(destination);
                })
    
            } catch (error) {
                console.log(error)
            }
        }
    },
    sheetToArray: function (sheet) {
        var result = [];
        var row;
        var rowNum;
        var colNum;
        var rangeAuto = xlsx.utils.decode_range(sheet['!ref'])
        var range = { s: { c: 0, r: 0 }, e: { c: (rangeAuto.e.c == 0) ? 0 : 2, r:  rangeAuto.e.r} };
        for(rowNum = range.s.r; rowNum <= range.e.r; rowNum++){
            row = [];
            for(colNum=range.s.c; colNum<=range.e.c; colNum++){
                var nextCell = sheet[
                    xlsx.utils.encode_cell({r: rowNum, c: colNum})
                ];
                if( typeof nextCell === 'undefined' ){
                    row.push(void 0);
                } else row.push(nextCell.w);
            }
            result.push(row);
        }
        return result;
    },
    readAndWriteExcel: async function (file, destination) {
        try {
            var destinationFile = destination.split('\\').slice(0,-1).toString().replace(/,/g,'\\');
            var propContent = '';
            var excelFile = xlsx.readFile(file)
            for (let i = 0; i < excelFile.SheetNames.length; i++) {
                worksheetName = excelFile.SheetNames[i];
                workSheet = excelFile.Sheets[worksheetName]
                var rowArray = Utils.sheetToArray(workSheet);
            
                for (let r = 0; r < rowArray.length; r++) {
                    var singleRow = rowArray[r];

                    // 0 column a, 1 column b etc....
                    if (singleRow[0] !== undefined && singleRow[2] !== undefined) {
                        singleRow[2] = singleRow[2].replace(/(\r\n|\n|\r)/gm,'');
                        singleRow[0] = singleRow[0].replace(/(\r\n|\n|\r)/gm,'');
                        propContent += `${singleRow[0]}=${singleRow[2]}\n`
                    }
                    
                    if (r == rowArray.length - 1 && propContent) {
                        try {
                            fs.writeFile(`${destinationFile}\\${worksheetName}.properties`, propContent, function(err) {
                                if(err) {
                                    return console.log(err);
                                }
                            });
                        } catch (err) {
                            console.log(err)
                        }
                        console.log(`The file: ${worksheetName} was saved!`);
                        propContent = '';
                    }
                }
            }
        } catch (err) {
            console.log(`Idk how but you fucked up ${err}`)
        }
    },
    copyAndReplaceFile: async function (file, destination, replacerObj) {
        try {
            await fs.copy(file, destination)
            fs.readFile(destination, function(err, data) {
                if (err) throw err
                JSZip.loadAsync(data).then(function (zip) {
                    var promises = []
                    var files = Object.keys(zip.files)
                    var regEx = /header|footer|document/g

                    for(i=0; i< files.length; i++){
                        if (files[i].match(regEx)) {
                            try {
                                promises.push(zip.file(files[i]).async("string"), files[i])
                            } catch (err) {
                                console.log(`Can't keep the promise ;( ${err}`)
                            }
                        }
                    }

                    Promise.all(promises).then((result)=> {
                        var regex = new RegExp(Object.keys(replacerObj).join("|"),"g")
                        for(i=0; i< result.length; i++){
                            if (i & 1) {
                                zip.file(result[i], result[i-1].replace(regex, (matched) => {
                                    return replacerObj[matched]
                                }))
                            }
                        }
                        zip.generateAsync({
                            type:"nodebuffer",
                            mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            compression: "DEFLATE",
                            compressionOptions: {level: 1}
                        }).then((content) => {
                            fs.writeFile(destination, content, function(err) {
                                if(err) {
                                    return console.log(err)
                                }
                                console.log(`The file: [${destination.split('\\').pop()}] was modified and saved!`)
                            }) 
                        })
                    })
                })
            })
        } catch (err) {
            console.log(err,"Something went wrong ehh")
        }
    },
    getFiles: async function (dir) {
        const subdirs = await readdir(dir)
        const files = await Promise.all(subdirs.map(async (subdir) => {
            const res = resolve(dir, subdir)
            return (await stat(res)).isDirectory() ? this.getFiles(res) : res
        }))
        return files.reduce((a, f) => a.concat(f), [])
    },
}

module.exports = Utils