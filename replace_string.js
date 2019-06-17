
var fs = require("fs")

let pending_files = []
let replace_map = {}

function replace_in_file(filename) {
    let code = fs.readFileSync(filename).toString()
    for (let key in replace_map) {
        code = code.split(key).join(replace_map[key])
    }
    // console.log("AFTER", code)
    fs.writeFileSync(filename, code)
}

for (var i = 2; i < process.argv.length; i++) {
    let arg = process.argv[i]
    if (arg.startsWith("--rep")) {
        i++
        let pair = process.argv[i].split("=")
        replace_map[pair[0]] = pair[1]
        console.log("replace", pair[0], "to", pair[1])
    } else {
        pending_files.push(arg)
    }
}


for (var i = 0; i < pending_files.length; i++) {
    let filename = pending_files[i]
    try {
        replace_in_file(filename)
        console.log("processed", filename)
    } catch (err) {
        console.error("process failed", filename)
    }
}
