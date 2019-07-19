
// node xlsx2pb.js path_to_proto_source_file.proto path_to_excel_file.xlsx --skip=0 --type=array --keyword=rows --sheet=Sheet1 --message=TestSheet1 --out=res/excel/out.json --out=res/excel/out.pb

let pbjs = require("protobufjs")
let fs = require("fs")
var xlsx = require("xlsx")
let scheme = {
    keyword: "rows", // 数组类型数据行字段固定命名
    keyId: "id",
    skip: 0,         // 忽略开头若干行
    packages: [],    // proto 文件内容对象化
    export: null,    // 需要转换的 message 类型对象
    rowType: null,   // 如果导出对象是数组默认嵌套类型 (method=array), 此为元素类型
    instance: null,  // 需要转换的 message 类型对象实例
    proto: "",       // 指定 proto 文件
    excel: "",       // 指定 excel 文件
    out: [],         // 输出文件
    workbook: null,  // excel 对象
    method: "array",
    sheet: "Sheet1", // 指定导出的 sheet
    message: "",     // 指定导出的 message 名称
    titles: [],
}

for (let i = 2; i < process.argv.length; i++) {
    let arg = process.argv[i]
    if (arg.endsWith(".proto")) {
        scheme.proto = arg
    } else if (arg.endsWith(".xlsx")) {
        scheme.excel = arg
    } else if (arg.startsWith("--type=")) {
        scheme.method = arg.substring("--type=".length)
    } else if (arg == "--type") {
        i++
        scheme.method = process.argv[i]
    } else if (arg.startsWith("--skip=")) {
        scheme.skip = parseInt(arg.substring("--skip=".length))
    } else if (arg == "--skip") {
        i++
        scheme.skip = parseInt(process.argv[i])
    } else if (arg.startsWith("--keyword=")) {
        scheme.keyword = arg.substring("--keyword=".length)
    } else if (arg == "--keyword") {
        i++
        scheme.keyword = process.argv[i]
    } else if (arg.startsWith("--sheet=")) {
        scheme.sheet = arg.substring("--sheet=".length)
    } else if (arg == "--sheet") {
        i++
        scheme.sheet = process.argv[i]
    } else if (arg.startsWith("--message=")) {
        scheme.message = arg.substring("--message=".length)
    } else if (arg == "--message") {
        i++
        scheme.message = process.argv[i]
    } else if (arg.startsWith("--out=")) {
        scheme.out.push(arg.substring("--out=".length))
    } else if (arg == "--out") {
        i++
        scheme.out.push(process.argv[i])
    }
}

console.log("proto:", scheme.proto)
console.log("excel:", scheme.excel)
console.log("sheet:", scheme.sheet)
console.log("message:", scheme.message)
console.log("method:", scheme.method)

// pbjs.parse(content, {keepComments: true})
let proto = pbjs.loadSync(scheme.proto)
for (let pb of proto.nestedArray) {
    // console.log("package:", pb.name)
    let package = {
        name: pb.name,
        types: {},
    }
    scheme.packages.push(package)
    for (let typename in pb) {
        let msgType = pb[typename]
        if (msgType instanceof pbjs.Type) {
            // msgType.create()
            // msgType.create().toJSON()
            package.types[typename] = msgType
            if (typename == scheme.message) {
                scheme.export = msgType;
                scheme.instance = msgType.create();
                // msgType.encode(scheme.instance).finish();
                // msgType.decode()
            }
            // console.log("type:", typename)
            for (let fieldKey in msgType.fields) {
                let field = msgType.fields[fieldKey]
                field.resolve()
                // console.log("field:", field.name, field.repeated ? "[]" : "*", field.type) // field.resolvedType
                // console.log("comment:", field.comment)
            }
        }
    }
}

// 检查表头名是否是有效的协议字段 
function check_field(type, fieldPath) {
    let field = type.fields[fieldPath[0]]
    if (field) {
        if (fieldPath.length == 1) {
            return true;
        }
        return check_field(field.resolvedType, fieldPath.slice(1));
    }
    return false;
}

function transform_primitive(typename, value, repeated) {
    switch (typename) {
        case "string":
            return "" + value
        case "bytes":
            return Buffer.from(value)
        case "int8": case "uint8":
        case "int16": case "uint16":
        case "int32": case "uint32":
        case "int64": case "uint64":
            if (repeated && typeof (value) == "string") {
                return value.split(";").map(it => parseInt(it))
            }
            return parseInt(value, 10)
        case "float":
        case "double":
            if (repeated && typeof (value) == "string") {
                return value.split(";").map(it => parseFloat(it))
            }
            return parseFloat(value)
        default:
            console.warn("unexpected typename:", typename)
            return parseFloat(value)
    }
}

function create_instance(type) {
    return type.create()
}

// instance: 实例
// fieldPath(string[]): 字段名层级
// value: 值 (来自excel单元格)
function assign_field(rowIndex, instance, type, fieldPath, value) {
    if (value == undefined) {
        return
    }

    let fieldName = fieldPath[0]
    let fieldTypeName = type.fields[fieldName].type
    let fieldValue = instance[fieldName]
    let fieldType = type.fields[fieldName].resolvedType;
    let fieldRepeated = type.fields[fieldName].repeated

    if (fieldType instanceof pbjs.Type) {
        if (fieldValue == null || fieldValue == undefined) {
            if (fieldRepeated) {
                console.error("never happen")
            } else {
                fieldValue = create_instance(fieldType)
                instance[fieldName] = fieldValue
                assign_field(rowIndex, fieldValue, fieldType, fieldPath.slice(1), value)
            }
        } else {
            if (fieldRepeated) {
                let fieldElement = null;
                if (!type.fields[fieldName].__lastRowIndex || type.fields[fieldName].__lastRowIndex != rowIndex) {
                    type.fields[fieldName].__lastRowIndex = rowIndex;
                    fieldElement = create_instance(fieldType);
                    fieldValue.push(fieldElement);
                } else {
                    fieldElement = fieldValue[fieldValue.length - 1];
                }
                assign_field(rowIndex, fieldElement, fieldType, fieldPath.slice(1), value)
            } else {
                assign_field(rowIndex, fieldValue, fieldType, fieldPath.slice(1), value)
            }
        }
    } else {
        // console.log(fieldRepeated, fieldTypeName)
        if (fieldValue == null || fieldValue == undefined) {
            if (fieldRepeated) {
                console.error("never happen")
            } else {
                fieldValue = transform_primitive(fieldTypeName, value)
                instance[fieldName] = fieldValue
            }
        } else {
            if (fieldRepeated) {
                let fieldElement = transform_primitive(fieldTypeName, value, fieldRepeated)
                if (fieldElement instanceof Array) {
                    fieldValue.push(...fieldElement)
                } else {
                    fieldValue.push(fieldElement)
                }
            } else {
                fieldValue = transform_primitive(fieldTypeName, value)
                // console.log(fieldName, fieldTypeName, fieldValue, value)
                instance[fieldName] = fieldValue
            }
        }
        // console.log(fieldName, ":", fieldTypeName, "=", fieldValue, ":", typeof (fieldValue), "**", value)
    }
}

console.log("export:", scheme.export.name)
let workbook = xlsx.readFile(scheme.excel)
scheme.workbook = workbook

for (let sheetIndex in workbook.SheetNames) {
    let sheetName = workbook.SheetNames[sheetIndex];
    if (sheetName == scheme.sheet) {
        let sheet = workbook.Sheets[sheetName];
        let range = xlsx.utils.decode_range(sheet["!ref"]);
        if (scheme.method == "array") {
            let rowObjects = []
            let rowObject = null
            scheme.instance[scheme.keyword] = rowObjects
            let rowType = scheme.export.fields[scheme.keyword].resolvedType // 数据实际类型
            scheme.rowType = rowType
            let startRowIndex = range.s.r + scheme.skip
            for (let rowIndex = startRowIndex; rowIndex <= range.e.r; rowIndex++) {
                if (rowIndex == startRowIndex) {
                    // 取字段名
                    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
                        let cell = sheet[xlsx.utils.encode_cell({ c: colIndex, r: rowIndex })]
                        if (cell && typeof (cell.v) == "string") {
                            let fieldPath = cell.v.split(".")
                            if (check_field(rowType, fieldPath)) {
                                // console.log("title:", cell.v)
                                scheme.titles.push(fieldPath)
                            } else {
                                scheme.titles.push(null)
                            }
                        } else {
                            scheme.titles.push(null)
                        }
                    }
                } else {
                    // 取值
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        let cell = sheet[xlsx.utils.encode_cell({ c: col, r: rowIndex })]
                        let title = scheme.titles[col]
                        if (title) {
                            if (title == scheme.keyId) {
                                if (!!cell && cell.v != "") {
                                    rowObject = rowType.create()
                                    rowObjects.push(rowObject)
                                }
                            }
                            if (rowObject) {
                                let cellValue = !!cell ? cell.v : undefined
                                assign_field(rowIndex, rowObject, rowType, title, cellValue)
                            }
                        }
                    }
                }
            }
        } else if (scheme.method == "object") {

        } else {

        }
    }
}

// output
for (let outfile of scheme.out) {
    if (outfile.endsWith(".json")) {
        fs.writeFileSync(outfile, JSON.stringify(scheme.instance))
    } else {
        let bytes = scheme.export.encode(scheme.instance).finish()
        fs.writeFileSync(outfile, bytes)

        // let obj = scheme.export.decode(bytes)
        // fs.writeFileSync(outfile + ".json", JSON.stringify(obj))
    }
}
