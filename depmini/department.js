const xlsx = require('xlsx')
const sql = require('alasql')
var config = {
    "姓名": {
        column: 'StaffName'
    },
    "年级": {
        column: 'Name',
        asColumn: 'AliasName'
    },
    "处组": {
        column: 'Name',
        asColumn: 'AliasName'
    },
    "序号": {
        column: 'OrderNum'
    }
};
module.exports = {
    ready: () => {
        console.log('to update departments');
        createTables();
        var dialog = require('electron').remote.dialog;
        /* show a file-open dialog and read the first selected file */
        var f = dialog.showOpenDialog({ properties: ['openFile'] });
        importFromExcel(f[0]);
    }
}

var insertIntoDep;
var createTables = () => {
    sql("create table DepartmentRel (OrderNum int, Name string, AliasName string, StaffName string)");
    insertIntoDep = sql.compile("insert into DepartmentRel (OrderNum, Name, AliasName, StaffName) values (?,?,?,?)");
}


var importFromExcel = (filePath) => {
    var workbook = xlsx.readFile(filePath);
    workbook.SheetNames.forEach(function (sheetName) {
        var result = xlsx.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if (result.length > 0) {
            var item = {};
            result.forEach(function(row){
                for (var col in row) {
                    if (!config[col])
                        continue;
                    item[config[col].column] = row[col];
                    if (config[col].asColumn)
                        item[config[col].asColumn] = col;
                }
                if (!isNaN(item.OrderNum))
                    sql("insert into DepartmentRel (OrderNum, Name, AliasName, StaffName) values (" + item.OrderNum + ", '" + item.Name + "', '" + item.AliasName + "', '" + item.StaffName + "')");
            });
            console.log(sql("select * from DepartmentRel"));
        }
    });
}

