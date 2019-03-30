var express = require('express');
var router = express.Router();
const excelToJson = require('convert-excel-to-json');
var json2xls = require('json2xls');
router.use(json2xls.middleware);
var fs = require('fs');
let categorized_byIndustry = [];
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
categorized_byIndustry = [];
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
  
// console.log(resultFirst.Sheet1[0]);

console.log(categorized_byIndustry)
  res.render('index', { title: categorized_byIndustry });
});

router.get('/download_category/:item', function(req, res, next) {

  
  let category = req.params.item;
  console.log(category);
  console.log(categorized_byIndustry); 
  let index = categorized_byIndustry.findIndex((item) => {
    return item.title === category;
  });    
  let jsonArray = [];
  console.log(index);
  
  categorized_byIndustry[index].codes.forEach(code => {
      let fileName = '';
      if(fs.existsSync(`kse/${code}.xls`)) {
        fileName = `kse/${code}.xls`;
      } else if(fs.existsSync(`kse/${code}.xlsx`)) {
        fileName = `kse/${code}.xlsx`;
      } else {
        fileName = '';
      }
      if(fileName !== '') {
        var resultFirst = excelToJson({
          sourceFile: fileName,
          header:{
            rows: 1
        }, columnToKey: {
          A: 'symbol',
          B: 'date',
          C: 'open',
          D: 'high',
          E: 'low',
          F: 'close',
          G: 'volume'
      }
      });
      console.log('line:83')
      console.log(resultFirst.Sheet1.length);
      // if(jsonArray === null) {
      //   jsonArray = resultFirst.Sheet1;
      // } else {
      //   jsonArray.push(resultFirst.Sheet1)
      // }
      for (let i = 0; i < resultFirst.Sheet1.length; i++) {
        jsonArray.push(resultFirst.Sheet1[i]);
      }
      console.log(jsonArray.length)   
      
    }
   
    });
    console.log(jsonArray.length)
    res.xls(`${categorized_byIndustry[index].title}.xlsx`, jsonArray);
    // console.log(jsonArray)
    // var xls = json2xls(jsonArray);
    // fs.writeFileSync(`${categorized_byIndustry[index].title}.xlsx`, xls, 'binary');
  // console.log(req.params.item);
  
});

module.exports = router;
