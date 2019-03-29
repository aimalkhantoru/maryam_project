var express = require('express');
var router = express.Router();
const excelToJson = require('convert-excel-to-json');

/* GET home page. */
router.get('/', function(req, res, next) {
  const resultXlx = excelToJson({
    sourceFile: 'kse.xlsx',
    header:{
      rows: 1
  },
  columnToKey: {
    A: 'code',
    B: 'symobl',
    C: 'id_code',
    D: 'industry',
    E: 'industry_code'
}
});
// console.log(resultXlx.Sheet1);
let result = resultXlx.Sheet1;
let categorized_byIndustry = [];
result.forEach(element => {
  let index = categorized_byIndustry.findIndex(function(item){
    return item.title === element.industry_code;
  });
  if(index === -1) {
    categorized_byIndustry.push({title: element.industry_code, codes: [element.symobl]})
  } else {
    categorized_byIndustry[index].codes.push(element.symobl)
  }
  
});
console.log(categorized_byIndustry)
  res.render('index', { title: resultXlx.sheet1 });
});

module.exports = router;
