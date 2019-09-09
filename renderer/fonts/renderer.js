const Excel = require('exceljs/modern.nodejs');
const { shell } = require('electron')

//window.addEventListener('DOMContentLoaded', () => {
//});

const runBtn = document.getElementById('runBtn');
const prepBtn = document.getElementById('prepBtn');

prepBtn.addEventListener('click', () => {
  prepBtn.classList.remove('bigBtn--visible');
  prepBtn.classList.add('hidden');
  runBtn.classList.remove('hidden');
  runBtn.classList.add('runBtn--visible');
  shell.openItem('source.xlsx');
});

runBtn.addEventListener('click', () => {
  dataReformat();
});

async function dataReformat() {
  animationStart();
  let workbook = new Excel.Workbook();
  workbook = await workbook.xlsx.readFile('source.xlsx');
  let plWorksheet = workbook.getWorksheet('PL');
  let maiWorksheet = workbook.getWorksheet('MAI');
  let mappingWorksheet = workbook.getWorksheet('Mapping');
  let maiRowId = 2
  let reference = '';
  let trailNumber = '';
  let mapping = [];
  plWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    let wnValue = plWorksheet.getRow(rowNumber).getCell('D').value
    let maValue = plWorksheet.getRow(rowNumber).getCell('E').value
    let prevMaValue = plWorksheet.getRow(rowNumber-1).getCell('E').value
    let prevWnValue = plWorksheet.getRow(rowNumber-1).getCell('D').value
    let amtValue = plWorksheet.getRow(rowNumber).getCell('F').value
    let ordinalNumber = plWorksheet.getRow(rowNumber).getCell('A').value
    let date = plWorksheet.getRow(rowNumber).getCell('B').value
    let desc = plWorksheet.getRow(rowNumber).getCell('C').value
     if(wnValue != null && maValue != null && wnValue != 'Konto WN' && maValue != 'Konto MA') {
      maiWorksheet.getRow(maiRowId).getCell('J').value = wnValue;
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValue;
      maiWorksheet.getRow(maiRowId).getCell('G').value = date;
      maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
      maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
      maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
      maiRowId++
      maiWorksheet.getRow(maiRowId).getCell('J').value = maValue;
      var amtValueFlip = -Math.abs(amtValue);
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValueFlip;
      maiWorksheet.getRow(maiRowId).getCell('G').value = date;
      maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
      maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
      maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
      maiRowId++
    }
    else if(wnValue != null && maValue == null) {
      if (prevWnValue == null && prevMaValue == null) {
        reference = wnValue;
        trailNumber = ordinalNumber;
      }
      else if (prevWnValue == 'Konto WN' && prevMaValue == 'Konto MA'){
        reference = wnValue;
        trailNumber = ordinalNumber
      }
      else {
        maiWorksheet.getRow(maiRowId).getCell('J').value = wnValue;
        maiWorksheet.getRow(maiRowId).getCell('L').value = amtValue;
        maiWorksheet.getRow(maiRowId).getCell('G').value = date;
        maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
        maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
        maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
        maiRowId++
      }
    }
    else if(wnValue == null && maValue != null) {
      maiWorksheet.getRow(maiRowId).getCell('J').value = maValue;
      var amtValueFlip = -Math.abs(amtValue);
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValueFlip;
      maiWorksheet.getRow(maiRowId).getCell('G').value = date;
      maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
      maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
      maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
      maiRowId++
    }
  });
  mappingWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    let globalCodeValue =  mappingWorksheet.getRow(rowNumber+1).getCell('A').value;
    let wnCode = mappingWorksheet.getRow(rowNumber+1).getCell('D').value;
    let maCode = mappingWorksheet.getRow(rowNumber+1).getCell('E').value;
    let dsc = mappingWorksheet.getRow(rowNumber+1).getCell('B').value;
    let obj = {}
    obj = { 
      ['wn']: wnCode,  
      ['ma']: maCode,   
      ['src']: globalCodeValue,
      ['dsc']: dsc
    }
    mapping.push(obj);
  })
  workbook.removeWorksheet(4);
  workbook.removeWorksheet(5);
  maiWorksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
    let plCodeValue = maiWorksheet.getRow(rowNumber+1).getCell('J').value
    let plCodeValString = plCodeValue
    let ob = mapping.find(code => code.src == plCodeValString);
    let valueCheck = maiWorksheet.getRow(rowNumber+1).getCell('L').value;
    if (typeof ob != 'undefined') {
      if (valueCheck > 0){
        if (ob.wn != null){
          maiWorksheet.getRow(rowNumber+1).getCell('B').value = ob.wn;
          maiWorksheet.getRow(rowNumber+1).getCell('K').value = ob.dsc;
        }
        else if (ob.wn == null){
          maiWorksheet.getRow(rowNumber+1).getCell('B').value = plCodeValue + " " + "is missing WN Code";
        }
      }
      else if (valueCheck < 0){
        if (ob.ma != null){
          maiWorksheet.getRow(rowNumber+1).getCell('B').value = ob.ma;
          maiWorksheet.getRow(rowNumber+1).getCell('K').value = ob.dsc;
        }
        else if (ob.ma == null){
          maiWorksheet.getRow(rowNumber+1).getCell('B').value = plCodeValue + " " + "is missing MA Code";
        }
      }
    }
    else if(typeof ob === 'undefined'){
      maiWorksheet.getRow(rowNumber+1).getCell('B').value = plCodeValue + " " + "was not found";
    }
  });
  workbook.xlsx.writeFile('new.xlsx');
  animationStop();
}

function animationStart() {
  document.getElementById('roller').classList.remove('hidden');
  document.getElementById('roller').classList.add('roller--visible');
  document.getElementById('fadescreen').classList.remove('hidden');
  document.getElementById('fadescreen').classList.add('fadescreen--visible');
  document.getElementById('dynamicMsg').classList.remove('container__dynamicMsg--visible');
  document.getElementById('dynamicMsg').classList.add('hidden');
  document.getElementById('fadescreen').style.transition = 'opacity 0.2s ease-out';
  document.getElementById('dynamicMsg').style.transition = 'all 0.2s ease-out';
}

function animationStop() {
  document.getElementById('roller').classList.remove('roller--visible');
  document.getElementById('roller').classList.add('hidden');
  document.getElementById('fadescreen').classList.remove('fadescreen--visible');
  document.getElementById('fadescreen').classList.add('hidden');
  document.getElementById('dynamicMsg').classList.remove('hidden');
  document.getElementById('dynamicMsg').classList.add('container__dynamicMsg--visible');
  document.getElementById('staticMsg').style.opacity = '0.5';
}
