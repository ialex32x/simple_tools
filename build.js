// node this filepath.json imagename version.go

var fs = require("fs")
var os = require("os")

var jsonfile = null
var buildname = "anonymouse"
var gofile = null
for (var i = 1; i < process.argv.length; i++) {
    if (process.argv[i] == "--image") {
        buildname = process.argv[i + 1]
        i++
        continue
    }
    if (process.argv[i] == "--json") {
        jsonfile = process.argv[i + 1]
        i++
        continue
    }
    if (process.argv[i] == "--go") {
        gofile = process.argv[i + 1]
        i++
        continue
    }
}

// console.log(jsonfile, buildname)
var build = {}
if (fs.existsSync(jsonfile)) {
    var filecontent = fs.readFileSync(jsonfile)
    build = JSON.parse(filecontent)
}
build.name = buildname || "anonymous"
build.build = (build.build || 0) + 1
build.time = new Date().toString()
build.username = os.userInfo().username
fs.writeFileSync(jsonfile, JSON.stringify(build, null, "\t"))

if (!!gofile) {
    fs.writeFileSync(gofile, `package base
const (
    Version = ${build.build}
    TimeTag = "${build.time}"
    Author  = "${build.username}"
)
`)
}
