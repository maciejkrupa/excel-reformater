const Excel = require('exceljs/modern.nodejs')
const workbook = new Excel.Workbook();

let maCode;
let wnCode;
let amtValue;
let ordinalNumber;
let trailNumber;
let date;
let desc;
let reference;
let maiRowId = "2";
let mapping = [];

function createMapObject () {
  let plWorksheet = workbook.getWorksheet('PL');
  let bMapWorksheet = workbook.getWorksheet('Mapping Broker');
  let sMapWorksheet = workbook.getWorksheet('Mapping Service');
  let mapId = plWorksheet.getRow(1).getCell('A').value;
  if (mapId === "MAI Insurance Brokers Poland Sp. z o.o. 2019/2020") {
    bMapWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      let srcCode = bMapWorksheet.getRow(rowNumber+1).getCell('A').value;
      let wnCode = bMapWorksheet.getRow(rowNumber+1).getCell('D').value;
      let maCode = bMapWorksheet.getRow(rowNumber+1).getCell('E').value;
      let dsc = bMapWorksheet.getRow(rowNumber+1).getCell('B').value;
      let object = { 
        ['wn']: wnCode,  
        ['ma']: maCode,   
        ['dsc']: dsc,
        ['src']: srcCode
      }
      mapping.push(object);
    })
  }
  else if (mapId === "MAI Service Sp. z o.o. 2019/2020") {
    sMapWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      let srcCode = sMapWorksheet.getRow(rowNumber+1).getCell('A').value;
      let wnCode = sMapWorksheet.getRow(rowNumber+1).getCell('D').value;
      let maCode = sMapWorksheet.getRow(rowNumber+1).getCell('E').value;
      let dsc = sMapWorksheet.getRow(rowNumber+1).getCell('B').value;
      let object = { 
        ['wn']: wnCode,  
        ['ma']: maCode,   
        ['dsc']: dsc,
        ['src']: srcCode
      }
      mapping.push(object);
    })
  }
  else {
    alert("No mapping available for your data. Please enter a correct name in cell A1 of the first Sheet.");
  };
};

function mapWnCodes(){
  let maiWorksheet = workbook.getWorksheet('MAI');
  let ob = mapping.find(code => code.src == wnCode);
  if (typeof ob != 'undefined') {
    if(ob.wn != null) {
      maiWorksheet.getRow(maiRowId).getCell('B').value = ob.wn;
      maiWorksheet.getRow(maiRowId).getCell('K').value = ob.dsc;
    }
    else if (ob.wn == null){
      maiWorksheet.getRow(maiRowId).getCell('B').value = wnCode + " " + "is missing WN Code";
    }
  }
  else if(typeof ob === 'undefined'){
    maiWorksheet.getRow(maiRowId).getCell('B').value = wnCode + " " + "was not found";
  };
};

function mapMaCodes(){
  let maiWorksheet = workbook.getWorksheet('MAI');
  let ob = mapping.find(code => code.src == maCode);
  if (typeof ob != 'undefined') {
    if(ob.ma != null) {
      maiWorksheet.getRow(maiRowId).getCell('B').value = ob.ma;
      maiWorksheet.getRow(maiRowId).getCell('K').value = ob.dsc;
    }
    else if (ob.ma == null){
    maiWorksheet.getRow(maiRowId).getCell('B').value = maCode + " " + "is missing MA Code";
    }
  }
  else if(typeof ob === 'undefined'){
    maiWorksheet.getRow(maiRowId).getCell('B').value = maCode + " " + "was not found";
  };
};

function setCellData (){
  let plWorksheet = workbook.getWorksheet('PL');
  let maiWorksheet = workbook.getWorksheet('MAI');
  let mapId = plWorksheet.getRow(1).getCell('A').value;
  if (mapId === "MAI Insurance Brokers Poland Sp. z o.o. 2019/2020") {
    maiWorksheet.getRow(maiRowId).getCell('C').value = 'PL40';
  }
  else if (mapId === "MAI Service Sp. z o.o. 2019/2020") {
    maiWorksheet.getRow(maiRowId).getCell('C').value = 'PL41';
  }
  maiWorksheet.getRow(maiRowId).getCell('G').value = date;
  maiWorksheet.getRow(maiRowId).getCell('I').value = desc;
  maiWorksheet.getRow(maiRowId).getCell('H').value = reference;
  maiWorksheet.getRow(maiRowId).getCell('A').value = parseFloat(trailNumber + '.' + ordinalNumber);
};


function formatRows(){
  let plWorksheet = workbook.getWorksheet('PL');
  let maiWorksheet = workbook.getWorksheet('MAI');
  plWorksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    let prevmaCode = plWorksheet.getRow(rowNumber-1).getCell('E').value;
    let prevwnCode = plWorksheet.getRow(rowNumber-1).getCell('D').value;
    wnCode = plWorksheet.getRow(rowNumber).getCell('D').value;
    wnCodeNext = plWorksheet.getRow(rowNumber+1).getCell('D').value;
    maCode = plWorksheet.getRow(rowNumber).getCell('E').value;
    maCodeNext = plWorksheet.getRow(rowNumber+1).getCell('E').value;
    amtValue = plWorksheet.getRow(rowNumber).getCell('F').value;
    ordinalNumber = plWorksheet.getRow(rowNumber).getCell('A').value;
    date = plWorksheet.getRow(rowNumber).getCell('B').value;
    desc = plWorksheet.getRow(rowNumber).getCell('C').value;
    if(wnCode != "" && wnCode != null && wnCode != 'Konto WN' && wnCode != 'Opis' && maCode != "" && maCode != null && maCode != 'Konto MA' && maCode != 'Konto WN') {
      maiWorksheet.getRow(maiRowId).getCell('J').value = wnCode;
      mapWnCodes();
      maiWorksheet.getRow(maiRowId).getCell('L').value = amtValue;
      setCellData();
      maiRowId++
      maiWorksheet.getRow(maiRowId).getCell('J').value = maCode;
      mapMaCodes();
      maiWorksheet.getRow(maiRowId).getCell('L').value = -amtValue;
      setCellData();
      maiRowId++
    }
    else if(wnCode != null && wnCode != "" && maCode == null || maCode === "") {   
      if ((prevwnCode == 'Konto WN' && prevmaCode == 'Konto MA') || (prevwnCode == null && prevmaCode == null) || (prevwnCode === '' && prevmaCode === '')) {
        reference = wnCode;
        trailNumber = ordinalNumber;
      }
      else {
        maiWorksheet.getRow(maiRowId).getCell('J').value = wnCode;
        mapWnCodes();
        maiWorksheet.getRow(maiRowId).getCell('L').value = amtValue;
        setCellData();
        maiRowId++
      }
    }
    else if (maCode != null && maCode != "" && wnCode == null || wnCode === "") {
      maiWorksheet.getRow(maiRowId).getCell('J').value = maCode;
      mapMaCodes();     
      maiWorksheet.getRow(maiRowId).getCell('L').value = -amtValue;
      setCellData();
      maiRowId++
    };
  });
};

let Reformat = function() {
  animation.runApp();
  warkbook = workbook.xlsx.readFile(srcFilePath)
  .then(function(){
    createMapObject();
    formatRows();
    workbook.xlsx.writeFile(newFilePath)
    .then(function(){
      animation.stopApp();
    }, reason => {
      animation.stopApp();
      alert("File was busy. Close it and run the App again.")
    });
  });
};

function exportFunction() {
  exports.Reformat = Reformat;
};

exportFunction();