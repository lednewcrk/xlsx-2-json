var _ = require('lodash')
var fsys = require('file-system')
var fs = require('fs')
var XLSX = require('xlsx')
var workbook = XLSX.readFile('sheet2.xlsx')
var sheet_name_list = workbook.SheetNames

function xlsxToJson() {
    sheet_name_list.forEach(function (y) {
        var worksheet = workbook.Sheets[y]
        var headers = {}
        var data = []
        for (z in worksheet) {
            if (z[0] === '!') continue
            //parse out the column, row, and value
            var tt = 0
            for (var i = 0; i < z.length; i++) {
                if (!isNaN(z[i])) {
                    tt = i
                    break
                }
            }
            var col = z.substring(0, tt)
            var row = parseInt(z.substring(tt))
            var value = worksheet[z].v

            //store header names
            if (row == 1 && value) {
                headers[col] = value
                continue
            }

            if (!data[row]) data[row] = {}
            data[row][headers[col]] = value
        }
        //drop those first two rows which are empty
        data.shift()
        data.shift()

        fsys.writeFile('./data.json', JSON.stringify(data), function (err) {})
    })
}

function dataToJson() {
    const rawdata = fs.readFileSync('data.json')
    const data = JSON.parse(rawdata)

    const mapped = data.map(({ ITEM: action, 'Quant.': qty, Valor: value }) => {
        return {
            action,
            qty,
            value: _.round(value, 2)
        }
    })

    // console.log(mapped)
    fsys.writeFile('./actions.json', JSON.stringify(mapped), function (err) {})
}

// xlsxToJson()
dataToJson()
