const xlsx = require('xlsx')
var departmentsRel = {};
module.exports = {
    ready: ()=>{
        console.log('to update departments');
        var dialog = require('electron').remote.dialog;
        /* show a file-open dialog and read the first selected file */
        var f = dialog.showOpenDialog({ properties: ['openFile'] });
        var workbook = xlsx.readFile(f[0]);
        var result = {};
        workbook.SheetNames.forEach(function (sheetName) {
            var roa = xlsx.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            if (roa.length > 0) {
                result[sheetName] = roa;
            }
        });
        console.log(result);
    }
}

