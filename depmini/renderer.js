const {department, spread} = require('./startup.js')
const $ = require('jquery');
$(()=>{
    $('.btn-update-departments').on('click', department.ready);
    $('.btn-import-excel').on('click', spread.importExcel)
    $('.btn-export-excel').on('click', spread.exportExcel)
})
