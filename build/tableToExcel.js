var $btdRc$filesaver = require("file-saver");
var $btdRc$exceljs = require("exceljs");


function $parcel$interopDefault(a) {
  return a && a.__esModule ? a.default : a;
}

function $parcel$defineInteropFlag(a) {
  Object.defineProperty(a, '__esModule', {value: true, configurable: true});
}

function $parcel$export(e, n, v, s) {
  Object.defineProperty(e, n, {get: v, set: s, enumerable: true, configurable: true});
}

$parcel$defineInteropFlag(module.exports);

$parcel$export(module.exports, "default", () => $25f7b829076aede2$export$2e2bcd8739ae039);
const $5dd49e5016284773$var$TTEParser = function() {
    let methods = {};
    /**
	 * Parse HTML table to excel worksheet
	 * @param {object} ws The worksheet object
	 * @param {HTML entity} table The table to be converted to excel sheet
	 */ methods.parseDomToTable = function(ws, table, opts) {
        const startrow = ws.rowCount + 1;
        let _r, _c, cs, rs, r, c;
        let rows = [
            ...table.rows
        ];
        let widths = table.getAttribute("data-cols-width");
        if (widths) widths = widths.split(",").map(function(item) {
            return parseInt(item);
        });
        let merges = [];
        for(_r = 0; _r < rows.length; ++_r){
            let row = rows[_r];
            r = _r + 1 + startrow; // Actual excel row number
            c = 1; // Actual excel col number
            if (row.getAttribute("data-exclude") === "true") {
                rows.splice(_r, 1);
                _r--;
                continue;
            }
            if (row.getAttribute("data-height")) {
                let exRow = ws.getRow(r);
                exRow.height = parseFloat(row.getAttribute("data-height"));
            }
            let tds = [
                ...row.children
            ];
            for(_c = 0; _c < tds.length; ++_c){
                let td = tds[_c];
                if (td.getAttribute("data-exclude") === "true") {
                    tds.splice(_c, 1);
                    _c--;
                    continue;
                }
                for(let _m = 0; _m < merges.length; ++_m){
                    var m = merges[_m];
                    if (m.s.c == c && m.s.r <= r && r <= m.e.r) {
                        c = m.e.c + 1;
                        _m = -1;
                    }
                }
                let exCell = ws.getCell(getColumnAddress(c, r));
                // calculate merges
                cs = parseInt(td.getAttribute("colspan")) || 1;
                rs = parseInt(td.getAttribute("rowspan")) || 1;
                if (cs > 1 || rs > 1) merges.push({
                    s: {
                        c: c,
                        r: r
                    },
                    e: {
                        c: c + cs - 1,
                        r: r + rs - 1
                    }
                });
                c += cs;
                exCell.value = getValue(td);
                if (!opts.autoStyle) {
                    let styles = getStylesDataAttr(td);
                    exCell.font = styles.font || null;
                    exCell.alignment = styles.alignment || null;
                    exCell.border = styles.border || null;
                    exCell.fill = styles.fill || null;
                    exCell.numFmt = styles.numFmt || null;
                }
            }
        }
        //Setting column width
        if (widths) widths.forEach((width, _i)=>{
            //у столбца будет максимальный width, в случае нескольких таблиц на листе
            if (!ws.columns[_i].width || ws.columns[_i].width < width) ws.columns[_i].width = width;
        });
        applyMerges(ws, merges);
        return ws;
    };
    /**
	 * To apply merges on the sheet
	 * @param {object} ws The worksheet object
	 * @param {object[]} merges array of merges
	 */ let applyMerges = function(ws, merges) {
        merges.forEach((m)=>{
            ws.mergeCells(getExcelColumnName(m.s.c) + m.s.r + ":" + getExcelColumnName(m.e.c) + m.e.r);
        });
    };
    /**
	 * Convert HTML to plain text
	 */ let htmldecode = function() {
        let entities = [
            [
                "nbsp",
                " "
            ],
            [
                "middot",
                "\xb7"
            ],
            [
                "quot",
                '"'
            ],
            [
                "apos",
                "'"
            ],
            [
                "gt",
                ">"
            ],
            [
                "lt",
                "<"
            ],
            [
                "amp",
                "&"
            ]
        ].map(function(x) {
            return [
                new RegExp("&" + x[0] + ";", "g"),
                x[1]
            ];
        });
        return function htmldecode(str) {
            let o = str.trim().replace(/\s+/g, " ").replace(/<\s*[bB][rR]\s*\/?>/g, "\n").replace(/<[^>]*>/g, "");
            for(let i = 0; i < entities.length; ++i)o = o.replace(entities[i][0], entities[i][1]);
            return o;
        };
    }();
    /**
	 * Takes a positive integer and returns the corresponding column name.
	 * @param {number} num  The positive integer to convert to a column name.
	 * @return {string}  The column name.
	 */ let getExcelColumnName = function(num) {
        for(var ret = "", a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26)ret = String.fromCharCode(parseInt(num % b / a) + 65) + ret;
        return ret;
    };
    let getColumnAddress = function(col, row) {
        return getExcelColumnName(col) + row;
    };
    /**
	 * Checks the data type specified and conerts the value to it.
	 * @param {HTML entity} td
	 */ let getValue = function(td) {
        let dataType = td.getAttribute("data-t");
        let rawVal = htmldecode(td.innerHTML);
        if (dataType) {
            let val;
            switch(dataType){
                case "n":
                    rawVal = rawVal.replace(/ /g, "").replace(",", ".");
                    val = Number(rawVal);
                    break;
                case "d":
                    let date = new Date(rawVal);
                    // To fix the timezone issue
                    val = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds()));
                    break;
                case "b":
                    val = rawVal.toLowerCase() === "true" ? true : rawVal.toLowerCase() === "false" ? false : Boolean(parseInt(rawVal));
                    break;
                default:
                    val = rawVal;
            }
            return val;
        } else if (td.getAttribute("data-hyperlink")) return {
            text: rawVal,
            hyperlink: td.getAttribute("data-hyperlink")
        };
        else if (td.getAttribute("data-error")) return {
            error: td.getAttribute("data-error")
        };
        return rawVal;
    };
    /**
	 * Prepares the style object for a cell using the data attributes
	 * @param {HTML entity} td
	 */ let getStylesDataAttr = function(td) {
        //Font attrs
        let font = {};
        if (td.getAttribute("data-f-name")) font.name = td.getAttribute("data-f-name");
        if (td.getAttribute("data-f-sz")) font.size = td.getAttribute("data-f-sz");
        if (td.getAttribute("data-f-color")) font.color = {
            argb: td.getAttribute("data-f-color")
        };
        if (td.getAttribute("data-f-bold") === "true") font.bold = true;
        if (td.getAttribute("data-f-italic") === "true") font.italic = true;
        if (td.getAttribute("data-f-underline") === "true") font.underline = true;
        if (td.getAttribute("data-f-strike") === "true") font.strike = true;
        // Alignment attrs
        let alignment = {};
        if (td.getAttribute("data-a-h")) alignment.horizontal = td.getAttribute("data-a-h");
        if (td.getAttribute("data-a-v")) alignment.vertical = td.getAttribute("data-a-v");
        if (td.getAttribute("data-a-wrap") === "true") alignment.wrapText = true;
        if (td.getAttribute("data-a-text-rotation")) alignment.textRotation = td.getAttribute("data-a-text-rotation");
        if (td.getAttribute("data-a-indent")) alignment.indent = td.getAttribute("data-a-indent");
        if (td.getAttribute("data-a-rtl") === "true") alignment.readingOrder = "rtl";
        // Border attrs
        let border = {
            top: {},
            left: {},
            bottom: {},
            right: {}
        };
        if (td.getAttribute("data-b-a-s")) {
            let style = td.getAttribute("data-b-a-s");
            border.top.style = style;
            border.left.style = style;
            border.bottom.style = style;
            border.right.style = style;
        }
        if (td.getAttribute("data-b-a-c")) {
            let color = {
                argb: td.getAttribute("data-b-a-c")
            };
            border.top.color = color;
            border.left.color = color;
            border.bottom.color = color;
            border.right.color = color;
        }
        if (td.getAttribute("data-b-t-s")) {
            border.top.style = td.getAttribute("data-b-t-s");
            if (td.getAttribute("data-b-t-c")) border.top.color = {
                argb: td.getAttribute("data-b-t-c")
            };
        }
        if (td.getAttribute("data-b-l-s")) {
            border.left.style = td.getAttribute("data-b-l-s");
            if (td.getAttribute("data-b-l-c")) border.left.color = {
                argb: td.getAttribute("data-b-t-c")
            };
        }
        if (td.getAttribute("data-b-b-s")) {
            border.bottom.style = td.getAttribute("data-b-b-s");
            if (td.getAttribute("data-b-b-c")) border.bottom.color = {
                argb: td.getAttribute("data-b-t-c")
            };
        }
        if (td.getAttribute("data-b-r-s")) {
            border.right.style = td.getAttribute("data-b-r-s");
            if (td.getAttribute("data-b-r-c")) border.right.color = {
                argb: td.getAttribute("data-b-t-c")
            };
        }
        //Fill
        let fill;
        if (td.getAttribute("data-fill-color")) fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: {
                argb: td.getAttribute("data-fill-color")
            }
        };
        //number format
        let numFmt;
        if (td.getAttribute("data-num-fmt")) numFmt = td.getAttribute("data-num-fmt");
        return {
            font: font,
            alignment: alignment,
            border: border,
            fill: fill,
            numFmt: numFmt
        };
    };
    return methods;
}();
var $5dd49e5016284773$export$2e2bcd8739ae039 = $5dd49e5016284773$var$TTEParser;




const $25f7b829076aede2$var$TableToExcel = function(Parser) {
    let methods = {};
    methods.initWorkBook = function() {
        let wb = new (0, ($parcel$interopDefault($btdRc$exceljs))).Workbook();
        return wb;
    };
    methods.initSheet = function(wb, sheetName) {
        let ws = wb.addWorksheet(sheetName);
        return ws;
    };
    methods.save = function(wb, fileName) {
        wb.xlsx.writeBuffer().then(function(buffer) {
            (0, ($parcel$interopDefault($btdRc$filesaver)))(new Blob([
                buffer
            ], {
                type: "application/octet-stream"
            }), fileName);
        });
    };
    methods.tableToSheet = function(wb, table, opts) {
        let ws = wb.getWorksheet(opts.sheet.name);
        if (ws === undefined) ws = this.initSheet(wb, opts.sheet.name);
        ws = Parser.parseDomToTable(ws, table, opts);
        return wb;
    };
    methods.tableToBook = function(table, opts) {
        let wb = this.initWorkBook();
        wb = this.tableToSheet(wb, table, opts);
        return wb;
    };
    methods.convert = function(table, opts = {}) {
        let defaultOpts = {
            name: "export.xlsx",
            autoStyle: false,
            sheet: {
                name: "Sheet 1"
            }
        };
        opts = {
            ...defaultOpts,
            ...opts
        };
        let wb = this.tableToBook(table, opts);
        this.save(wb, opts.name);
    };
    return methods;
}((0, $5dd49e5016284773$export$2e2bcd8739ae039));
var $25f7b829076aede2$export$2e2bcd8739ae039 = $25f7b829076aede2$var$TableToExcel;
window.TableToExcel = $25f7b829076aede2$var$TableToExcel;


//# sourceMappingURL=tableToExcel.js.map
