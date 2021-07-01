
const utils = require('./resourceUtils')
var type;

var replacerObj = {
    FIRMA01: "S.C. Some Name S.R.L.",
}

utils.getFiles("./sources").then(async (files)=> {
    var numOfFiles = {word:0,excel:0,properties:0,other:0}
    var allProperties = []
    for (let i = 0; i < files.length; i++) {
        let destination = files[i].replace(/sources/g, 'result')
        type = files[i].split(".").pop()
        fileName = files[i].split('\\').pop()

        if (type === "docx") {
            numOfFiles.word++
            await utils.copyAndReplaceFile(files[i], destination, replacerObj)
        } else if (type === "xlsx"){
            numOfFiles.excel++
            // for now please delete empty sheets....
            console.log(`Excel file: [${fileName}]`)
            utils.readAndWriteExcel(files[i], destination)
        } else if (type === "properties"){
            numOfFiles.properties++
            allProperties.push(files[i])
            console.log(`Properties files: [${fileName}]`)
        } else {
            numOfFiles.other++
            console.log(`File of type: ${type} is not supported!`)
        }
    }
    if (numOfFiles.properties > 0) {
        var ans = '';
        var type = '';
        while (ans.length < 2) {
            ans = await utils.askQuestion("Name of the excel ? ");
        }
        while (type.length < 2) {
            type = await utils.askQuestion("what do you want to do ? ");
        }
        utils.readAndWriteProperties(allProperties, `.\\result\\${ans}.xlsx`, type)
    }
    console.log(numOfFiles)
}).catch(err => console.error(err))