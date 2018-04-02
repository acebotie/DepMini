const xlsx = require('xlsx')
const sql = require('alasql')
module.exports = {
    importExcel: () => {
        var dialog = require('electron').remote.dialog;
        /* show a file-open dialog and read the first selected file */
        var f = dialog.showOpenDialog({ properties: ['openFile'] });
        importFromExcel(f[0]);
    },
    exportExcel: () => {
        console.log('to export a excel');
    }
}

var config = {
    "姓名": {
        column: 'Name'
    }
}
var importFromExcel = (filePath) => {
    console.log(sql("select * from DepartmentRel"));
    var workbook = xlsx.readFile(filePath);
    console.log(workbook);
    workbook.SheetNames.forEach(function (sheetName) {
        var result = xlsx.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if (result.length > 0) {
            console.log(result);
        }
    });
}