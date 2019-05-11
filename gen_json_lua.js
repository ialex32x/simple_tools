/*

*/

var fs = require("fs")
var path = require("path")

function get_arg_value(name, splitter) {
    var index = process.argv.findIndex(every => every == name)
    if (index >= 0) {
        var arg = process.argv[index + 1]
        if (arg) {
            return splitter ? arg.split(splitter) : arg
        }
    }
    return undefined
}

var json_filename = process.argv[2]
var json_basename = path.basename(json_filename).replace(path.extname(json_filename), "")
var json_dirname = path.dirname(json_filename)
var export_path = get_arg_value("--export")

if (export_path && !fs.existsSync(export_path)) {
    console.error("指定的导出路径不存在", export_path)
    return;
}

function translate_lua(json, json_key, level) {
    if (typeof(json) == "string") {
        return "\"" + json + "\""
    } else if (typeof(json) == "number") {
        return json + ""
    } else if (typeof(json) == "boolean") {
        // return json + "!!!"
        return (json ? "true" : "false") + "\""
    } else if (json instanceof Array) {
        var lines = "{"
        for (var key in json) {
            var value = json[key]
            lines += translate_lua(value, key, level + 1) 
            if (key < json.length - 1) {
                lines += ","
            }
        }
        return lines + "}"
    }
    var lines = "{\n"
    for (var key in json) {
        var value = json[key]
        lines += "[\"" + key + "\"] = " 
        if (value != undefined) {
            lines += translate_lua(value, key, level + 1) 
        } else {
            lines += "undefined"
        }
        lines += ",\n"
    }
    lines += "}"
    return lines
}

var json = JSON.parse(fs.readFileSync(json_filename))
var lua = "return " + translate_lua(json, null, 0)

var final_export_path = export_path 
                            ? path.join(export_path, json_basename) 
                            : path.join(json_dirname, json_basename)

console.log("export", json_basename)
fs.writeFileSync(final_export_path + ".lua", lua)
console.log("done.")