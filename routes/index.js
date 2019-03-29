var express = require('express');
var router = express.Router();
const excelToJson = require('convert-excel-to-json');

/* GET home page. */
router.get('/', function(req, res, next) {
  const result = excelToJson({
    sourceFile: 'kse.xlsx'
});
console.log(result.Sheet1[0]);
  res.render('index', { title: result.sheet1 });
});

module.exports = router;
