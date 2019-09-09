const Excel = require('exceljs/modern.nodejs');

window.addEventListener('DOMContentLoaded', () => {
  let el = document.getElementById('clickMe');
  if (el.addEventListener) {
    el.addEventListener('click', dataReformat, false);
  }
});

async function dataReformat() {
  animationStart();
  let workbook = new Excel.Workbook();
  workbook = await workbook.xlsx.readFile('data.xlsx');
  let plWorksheet = workbook.getWorksheet('PL');
  let maiWorksheet = workbook.getWorksheet('MAI');
  let mappingWorksheet = workbook.getWorksheet('Mapping');
  let maiRowId = 2
  let reference = '';
  let trailNumber = 'test';
  let mapping = [];
  plWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    let wnValue = plWorksheet.getRow(rowNumber).getCell('D').value
    let maValue = plWorksheet.getRow(rowNumber).getCell('E').value
    let prevMaValue = plWorksheet.getRow(rowNumber-1).getCell('E').value
    let prevWnValue = plWorksheet.getRow(rowNumber-1).getCell('D').value
    let amtString = plWorksheet.getRow(rowNumber).getCell('F').value
    var amtValue = parseFloat(amtString);
    let ordinalNumber = plWorksheet.getRow(rowNumber).getCell('A').value
    let date = plWorksheet.getRow(rowNumber).getCell('B').value
    let desc = plWorksheet.getRow(rowNumber).getCell('C').value
     if(wnValue != null && maValue != null && wnValue != 'Konto WN' && maValue != 'Konto MA') {
      maiWorksheet.getRow(maiRowId).getCell('J').value = wnValue;
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValue;
      maiWorksheet.getRow(maiRowId).getCell('G').value = date;
      maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
      maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
      //maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
      maiRowId++
      maiWorksheet.getRow(maiRowId).getCell('J').value = maValue;
      var amtValueFlip = -Math.abs(amtValue);
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValueFlip;
      maiWorksheet.getRow(maiRowId).getCell('G').value = date;
      maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
      maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
      //maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
      maiRowId++
    }
    // else if(wnValue != null && maValue == null) {
    //   //if (prevWnValue == null && prevMaValue == null) {
    //     // reference = wnValue;
    //     // trailNumber = ordinalNumber;
    //   //}
    //   //if (prevWnValue == 'Konto WN' && prevMaValue == 'Konto MA'){
    //     // reference = wnValue;
    //     // trailNumber = ordinalNumber
    //   //}
    //   //else {
    //     maiWorksheet.getRow(maiRowId).getCell('J').value = wnValue;
    //     maiWorksheet.getRow(maiRowId).getCell('L').value = amtValue;
    //     maiWorksheet.getRow(maiRowId).getCell('G').value = date;
    //     maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
    //     maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
    //    //maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
    //     maiRowId++
    //   //}
    // }
    else if(wnValue == null && maValue != null) {
      maiWorksheet.getRow(maiRowId).getCell('J').value = maValue;
      var amtValueFlip = -Math.abs(amtValue);
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValueFlip;
      maiWorksheet.getRow(maiRowId).getCell('G').value = date;
      maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
      maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
      //maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
      maiRowId++
    }
  });
  // mappingWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
  //   let globalCodeValue =  mappingWorksheet.getRow(rowNumber+1).getCell('A').value;
  //   let wnCode = mappingWorksheet.getRow(rowNumber+1).getCell('D').value;
  //   let maCode = mappingWorksheet.getRow(rowNumber+1).getCell('E').value;
  //   let dsc = mappingWorksheet.getRow(rowNumber+1).getCell('B').value;
  //   let obj = {}
  //   obj = { 
  //     ['wn']: wnCode,  
  //     ['ma']: maCode,   
  //     ['src']: globalCodeValue,
  //     ['dsc']: dsc
  //   }
  //   mapping.push(obj);
  // })
  // maiWorksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
  //   let plCodeValue = maiWorksheet.getRow(rowNumber+1).getCell('J').value
  //   let plCodeValString = plCodeValue
  //   let ob = mapping.find(code => code.src == plCodeValString);
  //   let valueCheck = maiWorksheet.getRow(rowNumber+1).getCell('L').value;
  //   if (typeof ob != 'undefined') {
  //     if (valueCheck > 0){
  //       if (ob.wn != null){
  //         maiWorksheet.getRow(rowNumber+1).getCell('B').value = ob.wn;
  //       }
  //       else if (ob.wn == null){
  //         maiWorksheet.getRow(rowNumber+1).getCell('B').value = plCodeValue + " " + "is missing WN Code";
  //       }
  //     }
  //     else if (valueCheck < 0){
  //       if (ob.ma != null){
  //         maiWorksheet.getRow(rowNumber+1).getCell('B').value = ob.ma;
  //       }
  //       else if (ob.ma == null){
  //         maiWorksheet.getRow(rowNumber+1).getCell('B').value = plCodeValue + " " + "is missing MA Code";
  //       }
  //     }
  //   }
  //   else if(typeof ob === 'undefined'){
  //     maiWorksheet.getRow(rowNumber+1).getCell('B').value = plCodeValue + " " + "was not found";
  //   }
  // });
  workbook.xlsx.writeFile('test.xlsx');
  animationStop();
}

function animationStart() {
  document.getElementById('roller').classList.remove('hidden');
  document.getElementById('roller').classList.add('visible_roller');
  document.getElementById('blackout').classList.remove('hidden');
  document.getElementById('blackout').classList.add('visible_blackout');
  document.getElementById('msg1').classList.remove('visible_msg');
  document.getElementById('msg1').classList.add('hidden');
  document.getElementById('blackout').style.transition = 'opacity 0.2s ease-out';
  document.getElementById('msg1').style.transition = 'all 0.2s ease-out';
}

function animationStop() {
  document.getElementById('roller').classList.remove('visible_roller');
  document.getElementById('roller').classList.add('hidden');
  document.getElementById('blackout').classList.remove('visible_blackout');
  document.getElementById('blackout').classList.add('hidden');
  document.getElementById('msg1').classList.remove('hidden');
  document.getElementById('msg1').classList.add('visible_msg');
  document.getElementById('msg0').style.opacity = '0.5';
}
