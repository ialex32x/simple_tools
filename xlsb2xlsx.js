#!/usr/bin/env node

var xlsx = require("xlsx")
var fs = require("fs")
var path = require("path")

var curPath = path.resolve(__dirname) 
var files = fs.readdirSync(curPath)

console.log("js " + curPath)
var file_extension = ".xlsb"
for (var index in files) {
    var filename = files[index]
    if (filename.endsWith(file_extension)) {
        var inpath = path.join(curPath, filename)
        var outpath = path.join(curPath, filename.substring(0, filename.length - file_extension.length) + ".xlsx")

        console.log("convert: " + inpath)
        var wb = xlsx.readFile(inpath)
        xlsx.writeFile(wb, outpath)
    }
}
