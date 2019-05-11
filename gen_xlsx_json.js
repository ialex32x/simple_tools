/*

--sheets 表格名字1,表格名字2
--sheets 表格名字1:,表格名字2:重命名
    指定要导出的表格
--human 
    生成可阅读的json
--lua 
    同时生成lua
--skip N 
    跳过前N行
--export path
    指定导出结果目录,如果没有指定,那么导出到源文件目录
--id Id
    指定用于索引的列 (不指定时当做数组处理)
--typed 字段1 字段2:=int[],字段3:=float[]
    指定某些字段的类型, 可指定的类型为 int float string int[] float[] int:int int:float 
    类型不显式指定的列, 由解释器自动推导

 */
//"args": ["../../excel/role_test2.xlsx", "--id", "Id", "--sheets", "Sheet4:", "--human", "--lua", "--skip", "1", "--export", "../../export", "--typed", "Life Attack MoveSpeed AttackSpeed Defense HitRate DodgeRate LifeGet CritRate CritValue CdReduce:=float[],inline:=int:int"],

var fs = require("fs")
var xlsx = require("xlsx")
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

var xlsx_filename = process.argv[2]
var xlsx_basename = path.basename(xlsx_filename).replace(path.extname(xlsx_filename), "")
var xlsx_dirname = path.dirname(xlsx_filename)
var xlsx_sheets_raw = get_arg_value("--sheets", ",")
var xlsx_sheets = []
var export_path = get_arg_value("--export")
var index_by_key = get_arg_value("--id")
var types_def = get_arg_value("--typed", ",")
var skip_rows_str = get_arg_value("--skip")
var skip_rows = 0

if (skip_rows_str) {
    skip_rows = parseInt(skip_rows_str)
}

if (export_path && !fs.existsSync(export_path)) {
    console.error("指定的导出路径不存在", export_path)
    return;
}

var is_human = process.argv.findIndex(every => every == "--human") >= 0
var is_no_json = process.argv.findIndex(every => every == "--no-json") >= 0
var is_export_lua = process.argv.findIndex(every => every == "--lua") >= 0

var sheet_name_map = {}
for (var index in xlsx_sheets_raw) {
    var kv = xlsx_sheets_raw[index].split(":")
    if (kv.length == 2) {
        sheet_name_map[kv[0]] = kv[1]
        console.log(`重命名 ${kv[0]} => ${kv[1]}`)
    }
    xlsx_sheets.push(kv[0])
}

function get_filename(part1, part2) {
    if (sheet_name_map[part2] != undefined) {
        part2 = sheet_name_map[part2]
    }
    return (part1 == part2 || part2 == "" || !part2) ? part1 : (part1 + "_" + part2)
}

var title_type_map = {}
if (types_def) {
    for (var index in types_def) {
        var titles_type = types_def[index].split(":=")
        var titles = titles_type[0].split(" ")
        for (var title_index in titles) {
            title_type_map[titles[title_index]] = titles_type[1]
            // console.log(`指定字段类型 ${titles[title_index]} = ${titles_type[1]}`)
        }
    }
}

function translate_value(title, value, last, luafs) {
    var type = title && title_type_map[title]
    // console.log(`try parse ${title} as ${type}`)
    switch (type) {
        case "float[]": {
            var values = Array.from(value.split(","), (v, k) => parseFloat(v)) 
            if (luafs) {
                fs.writeSync(luafs, "{")
                for (var index in values) {
                    fs.writeSync(luafs, index == values.length - 1 ? `${values[index]}` : `${values[index]},`)
                }
                fs.writeSync(luafs, last ? "}" : "},")
            }
            return values
        }
        case "int[]": {
            var values = Array.from(value.split(","), (v, k) => parseInt(v)) 
            if (luafs) {
                fs.writeSync(luafs, "{")
                for (var index in values) {
                    fs.writeSync(luafs, index == values.length - 1 ? `${values[index]}` : `${values[index]},`)
                }
                fs.writeSync(luafs, last ? "}" : "},")
            }
            return values
        }
        case "int:int": {
            var values = value.split(",")
            var map = {}
            for (var index in values) {
                var value_kv = Array.from(values[index].split(":"), (v, k) => parseInt(v))
                map[value_kv[0]] = value_kv[1]
            }
            if (luafs) {
                fs.writeSync(luafs, "{")
                var index = 0
                for (var mapkey in map) {
                    fs.writeSync(luafs, index++ == values.length - 1 ? `[${mapkey}]=${map[mapkey]}` : `[${mapkey}]=${map[mapkey]},`)
                }
                fs.writeSync(luafs, last ? "}" : "},")
            }
            return map
        }
        case "int:float": {
            var values = value.split(",")
            var map = {}
            for (var index in values) {
                var value_kv = Array.from(values[index].split(":"), (v, k) => k % 2 ? parseInt(v) : parseFloat(v))
                map[value_kv[0]] = value_kv[1]
            }
            if (luafs) {
                fs.writeSync(luafs, "{")
                var index = 0
                for (var mapkey in map) {
                    fs.writeSync(luafs, index++ == values.length - 1 ? `[${mapkey}]=${map[mapkey]}` : `[${mapkey}]=${map[mapkey]},`)
                }
                fs.writeSync(luafs, last ? "}" : "},")
            }
            return map
        }
        case "int": {
            if (luafs) {
                fs.writeSync(luafs, last ? `${value}` : `${value},`)
            }
            return typeof(value) == "number" ? value : parseInt(`${value}`)
        }
        case "float": {
            if (luafs) {
                fs.writeSync(luafs, last ? `${value}` : `${value},`)
            }
            return typeof(value) == "number" ? value : parseFloat(`${value}`)
        }
        case "string": {
            if (luafs) {
                fs.writeSync(luafs, last ? `"${value}"` : `"${value}",`)
            }
            return `${value}`
        }
        default: break
    }
    if (luafs) {
        if (typeof(value) == "string") {
            fs.writeSync(luafs, last ? `"${value}"` : `"${value}",`)
        } else {
            fs.writeSync(luafs, last ? `${value}` : `${value},`)
        }
    }
    return value
}

var wb = xlsx.readFile(xlsx_filename)

for (var sheetIndex in wb.SheetNames) {
    var sheetName = wb.SheetNames[sheetIndex]
    if (xlsx_sheets) {
        if (xlsx_sheets.findIndex(every => {
            // console.log("?", every, sheetName)
            return every == sheetName
        }) < 0) {
            continue
        }
    }
    var final_export_path = export_path 
                            ? path.join(export_path, get_filename(xlsx_basename, sheetName)) 
                            : path.join(xlsx_dirname, get_filename(xlsx_basename, sheetName))

    console.log(`export sheet: ${xlsx_filename}:${sheetName}`)

    var luafs = is_export_lua && fs.openSync(final_export_path + ".lua", "w")
    var sheet = wb.Sheets[sheetName]
    var jsheet = {
        name: sheetName, 
        title: [],
        types: title_type_map,
        rows: 0
    }
    if (luafs) {
        fs.writeSync(luafs, "return {\n")
        fs.writeSync(luafs, `["name"]="${sheetName}",\n`)
        fs.writeSync(luafs, `["values"]={\n`)
    }
    var range = xlsx.utils.decode_range(sheet["!ref"])
    var titleCheck = []
    var titleRaw = []
    var titleRead = true
    var key_index = -1
    var values = undefined
    for (var row = range.s.r; row <= range.e.r; row++) {
        if (skip_rows-- > 0) {
            continue
        }

        if (titleRead) {
            titleRead = false
            for (var col = range.s.c; col <= range.e.c; col++) {
                var cell = sheet[xlsx.utils.encode_cell({c: col, r: row})]
                if (cell) {
                    if (typeof(cell.v) == "string" && cell.v != "") {
                        jsheet.title.push(cell.v)
                        titleCheck.push(true)
                        titleRaw.push(cell.v)
                        if (index_by_key && cell.v == index_by_key) {
                            key_index = col
                            jsheet.index = cell.v
                        } 
                    } else {
                        titleCheck.push(false)
                        titleRaw.push(false)
                    }
                } else {
                    titleRaw.push(false)
                    titleCheck.push(false)
                }
            }
            if (key_index >= 0) {
                values = {}
            } else {
                values = []
            }
            // console.log(`title: ${jsheet.title}`)
        } else {
            var jrow = []

            if (key_index >= 0) {
                var cell = sheet[xlsx.utils.encode_cell({c: key_index, r: row})]
                if (luafs) {
                    if (typeof(cell.v) == "string") {
                        fs.writeSync(luafs, `["${cell.v}"]=`)
                    } else {
                        fs.writeSync(luafs, `[${cell.v}]=`)
                    }
                }
                values[cell.v] = jrow
            } else {
                values.push(jrow)
            }

            if (luafs) {
                fs.writeSync(luafs, "{")
            }

            var col_counter = 0
            for (var col = range.s.c; col <= range.e.c; col++) {
                if (titleCheck[col]) {
                    var cell = sheet[xlsx.utils.encode_cell({c: col, r: row})]
                    var cell_value = cell != undefined && cell != null && cell.v
            
                    // console.log(`${titleRaw[col]}: col ${col} row ${row} = ${cell_value}`)
                    jrow.push(translate_value(
                        titleRaw[col], 
                        cell_value, 
                        col_counter == jsheet.title.length - 1, 
                        luafs))
                    col_counter++
                }
            }
            
            if (luafs) {
                fs.writeSync(luafs, row == range.e.r ? "}\n" : "},\n")
            }
            jsheet.rows += 1
        } // end if row
    } // end for row
    jsheet.values = values
    if (luafs) {
        fs.writeSync(luafs, `},\n`)
        fs.writeSync(luafs, `["title"]={`)
        for (var col in jsheet.title) {
            if (col < jsheet.title.length - 1) {
                fs.writeSync(luafs, `"${jsheet.title[col]}",`)
            } else {
                fs.writeSync(luafs, `"${jsheet.title[col]}"`)
            }
        }
        fs.writeSync(luafs, `},\n`)
        fs.writeSync(luafs, `["index"]=${jsheet.index},\n`)
        fs.writeSync(luafs, `["rows"]=${jsheet.rows}\n`)
        
        fs.writeSync(luafs, "}\n")
    }
    if (!is_no_json) {
        fs.writeFileSync(final_export_path + ".json", JSON.stringify(jsheet, null, is_human ? "\t" : null))
    }
    fs.closeSync(luafs)
}
console.log("done.")