var excelbuilder = require('msexcel-builder');

function excelize(obj, path, file, sheet, cb){
    var titles = Object.keys(obj[0]);
    var width = titles.length;
    var height = obj.length;
    console.log([titles,width,height]);
    var workbook = excelbuilder.createWorkbook(path, file);
    var sht = workbook.createSheet(sheet, width, height+1);
    var i, j, k;

    for (i = 1; i <= width; i++){
        sht.set(i, 1, titles[i-1]);
        sht.font(i, 1, {bold:'true'});
    }
    for (j = 1; j <= height; j++) {
        for (i = 1; i <= width; i++){
            k = titles[i-1];
            sht.set(i, j+1, obj[j-1][k]);
        }
    }
    workbook.save(cb);
}

module.exports = excelize;