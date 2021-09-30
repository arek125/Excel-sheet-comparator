
var wb1,wb2
//var XLSX = require("sheetjs-style")
import XLSX from 'sheetjs-style'

window.addEventListener('DOMContentLoaded', () => {

    var excel1 = document.getElementById("excel1")
    var excel2 = document.getElementById("excel2")
    var key = document.getElementById("mainkey")
    var sheetName1 = document.getElementById("excel1sheetname")
    var sheetName2 = document.getElementById("excel2sheetname")
    var mainForm = document.getElementById("mainForm")
    var compareButton = document.querySelector('#compere')
    

    compareButton.addEventListener('click',async () => {
        console.log(excel1.files)
        if(excel1.value != "" && excel2.value != "" && key.value && sheetName1.value && sheetName2.value) {
            // let arr1 = await parseExcel2(excel1.files[0].path,sheetName1.value)
            // let arr2 = await parseExcel2(excel2.files[0].path,sheetName2.value)
            try{
            compareButton.setAttribute("aria-busy", "true");
            let arr1 = parseWorkbook(wb1,sheetName1.value)
            let arr2 = parseWorkbook(wb2,sheetName2.value)
            console.log(arr1,arr2)
            let shared = arr1.filter(o1 => arr2.some(o2 => o1[key.value] === o2[key.value]));
            let notShared = arr1.filter(o1 => !arr2.some(o2 => o1[key.value] === o2[key.value]));
            console.log(shared,notShared)
            let diffs = []
            if(arr1.length && arr2.length){
              if(!arr1[0].hasOwnProperty(key.value) || !arr2[0].hasOwnProperty(key.value)){
                setInfo(key.value +"Key not exist !", true)
                return
              }
              var filesHeader = [excel1.files[0].name]
              for(var i = 1;i<Object.keys(arr1[0]).length;i++)filesHeader.push("")
              filesHeader.push(excel2.files[0].name)
              diffs.push(filesHeader)
              diffs.push(Object.keys(arr1[0]).concat(Object.keys(arr2[0])))
            }
            else {
              setInfo("Empty sheet !", true)
              return
            }
            for( let el of shared){
                let i1 = arr1.findIndex(x=>x[key.value] == el[key.value])
                let i2 = arr2.findIndex(x=>x[key.value] == el[key.value])
                if(!shallowEqual(arr1[i1], arr2[i2])) {
                    // let newObj={}
                    // for (const [key, value] of Object.entries(arr1[i1])) {
                    //     newObj[key +'-'+ excel1.files[0].name] = value != arr2[i2][key] ? value+"*(": value
                    // }
                    // for (const [key, value] of Object.entries(arr2[i2])) {
                    //     newObj[key +'-'+ excel2.files[0].name] = value != arr1[i1][key] ? value+"*(": value
                    // }
                    // diffs.push(newObj)
                    let newArr=[]
                    for (const [key, value] of Object.entries(arr1[i1])) {
                        value.toString().replace(/-/g,"") != arr2[i2][key].toString().replace(/-/g,"") ? newArr.push(value+"###"): newArr.push(value.toString())
                    }
                    for (const [key, value] of Object.entries(arr2[i2])) {
                        value.toString().replace(/-/g,"") != arr1[i1][key].toString().replace(/-/g,"") ? newArr.push(value+"###"): newArr.push(value.toString())
                    }
                    diffs.push(newArr)
                }
            }
            for( let el of notShared){
              let i1 = arr1.findIndex(x=>x[key.value] == el[key.value])
              let i2 = arr2.findIndex(x=>x[key.value] == el[key.value])
              let newArr=[]
              if(i1 != -1)
                for (const [key, value] of Object.entries(arr1[i1])) {
                  newArr.push(value+"###")
                }
              else
                for (const i of Object.entries(arr1[0])) {
                  newArr.push("")
                }
              if(i2 != -1)
                for (const [key, value] of Object.entries(arr2[i2])) {
                  newArr.push(value+"###")
                }
              else
                for (const i of Object.entries(arr2[0])) {
                  newArr.push("")
                }
              diffs.push(newArr)
            }
            console.log(diffs)
            var wb = XLSX.utils.book_new();
            wb.SheetNames.push("Diff");
            wb.SheetNames.push(excel1.files[0].name);
            wb.Sheets[excel1.files[0].name] = XLSX.utils.json_to_sheet(arr1,{skipHeader: false});
            wb.SheetNames.push(excel2.files[0].name);
            wb.Sheets[excel2.files[0].name] = XLSX.utils.json_to_sheet(arr2,{skipHeader: false});
            //var ws = XLSX.utils.json_to_sheet(diffs,{skipHeader: false});
            var ws = XLSX.utils.aoa_to_sheet(diffs);
            ws['!merges'] = [{s: {c: 0, r:0 }, e: {c:Object.keys(arr1[0]).length-1, r:0}},{s: {c: Object.keys(arr1[0]).length, r:0 }, e: {c:Object.keys(arr1[0]).length+Object.keys(arr2[0]).length-1, r:0}}];
            for( let [ri, el] of diffs.entries()){
              if(ws[cL(Object.keys(arr1[0]).length)+(ri+1)] != "undefined")
                ws[cL(Object.keys(arr1[0]).length)+(ri+1)].s = {
                  border: { left: {style: "thick", color: "000000"}}
                };
              Object.entries(el).forEach(([key, value], oi) => {
                //console.log(ws[cL(oi)+(ri+2)])
                if(typeof ws[cL(oi)+(ri+2)] != "undefined" && ws[cL(oi)+(ri+2)].v && ws[cL(oi)+(ri+2)].v.toString().includes("###")){
                  ws[cL(oi)+(ri+2)].s = { fill: {patternType: "solid", fgColor: { rgb: "d07629" }}, }
                  ws[cL(oi)+(ri+2)].v = ws[cL(oi)+(ri+2)].v.replace( /###/, "");
                }
              });
            }
            wb.Sheets["Diff"] = ws
            XLSX.writeFile(wb, "Result.xlsx");
          }catch(e){
            console.log(e)
            setInfo(e.message, true)
          }finally{
            compareButton.setAttribute("aria-busy", "false");
          }
        }else setInfo("Choose files and fill all fields !", true)
      
    })
    excel1.onchange = async function() {
      wb1 = await readExcelFile(this.files[0])
      fillSelect(sheetName1,wb1.SheetNames)
    };
    excel2.onchange = async function() {
      wb2 = await readExcelFile(this.files[0])
      fillSelect(sheetName2,wb2.SheetNames)
    };
  })

function parseWorkbook(wb,sheetName){
  return XLSX.utils.sheet_to_row_object_array(wb.Sheets[sheetName],{defval:""})
}
function removeOptions(selectElement) {
  var i, L = selectElement.options.length - 1;
  for(i = L; i >= 0; i--) {
     selectElement.remove(i);
  }
}
function fillSelect(select,optionsArr){
  removeOptions(select);
  for(let el of optionsArr){
    var opt = document.createElement('option');
    opt.value = el;
    opt.innerHTML = el;
    select.appendChild(opt);
  }
  select.value = optionsArr[0]
}
async function readExcelFile(file) {
  return new Promise(function(resolve,reject){
      var reader = new FileReader();
      //reader.readAsDataURL(files[0]);
      reader.onload = function(e) {
          var data = e?e.target.result:reader.content;
          resolve(XLSX.read(data, {
              type: 'binary',
              cellDates: true
          }));
      }
      reader.onerror = function () {
          reject(reader.error.message);
      };
      reader.readAsBinaryString(file)
  })
}

function shallowEqual(object1, object2) {
    const keys1 = Object.keys(object1);
    const keys2 = Object.keys(object2);
  
    if (keys1.length !== keys2.length) {
      return false;
    }
  
    for (let key of keys1) {
      object1[key] = object1[key].toString().replace(/-/g,"")
      object2[key] = object2[key].toString().replace(/-/g,"")
      if (object1[key] !== object2[key]) {
        return false;
      }
    }
  
    return true;
  }

function cL(n) {//column letter
    var ordA = 'a'.charCodeAt(0);
    var ordZ = 'z'.charCodeAt(0);
    var len = ordZ - ordA + 1;
    var s = "";
    while(n >= 0) {
        s = String.fromCharCode(n % len + ordA) + s;
        n = Math.floor(n / len) - 1;
    }
    return s.toUpperCase();
}


function setInfo(info, err){
  if(!err)document.getElementById("output").innerHTML = "<ins>"+info+"</ins>";
  else document.getElementById("output").innerHTML = "<mark>"+info+"</mark>";
}


if (FileReader.prototype.readAsBinaryString === undefined) {
  FileReader.prototype.readAsBinaryString = function (fileData) {
      var binary = "";
      var pt = this;
      var reader = new FileReader();
      reader.onload = function (e) {
          var bytes = new Uint8Array(reader.result);
          var length = bytes.byteLength;
          for (var i = 0; i < length; i++) {
              binary += String.fromCharCode(bytes[i]);
          }
          //pt.result  - readonly so assign content to another property
          pt.content = binary;
          pt.onload(); // thanks to @Denis comment
      }
      reader.readAsArrayBuffer(fileData);
  }
}
// if (typeof Object.assign != 'function') {
//   Object.assign = function(target, varArgs) { // .length of function is 2
//     'use strict';
//     if (target == null) { // TypeError if undefined or null
//       throw new TypeError('Cannot convert undefined or null to object');
//     }

//     var to = Object(target);

//     for (var index = 1; index < arguments.length; index++) {
//       var nextSource = arguments[index];

//       if (nextSource != null) { // Skip over if undefined or null
//         for (var nextKey in nextSource) {
//           // Avoid bugs when hasOwnProperty is shadowed
//           if (Object.prototype.hasOwnProperty.call(nextSource, nextKey)) {
//             to[nextKey] = nextSource[nextKey];
//           }
//         }
//       }
//     }
//     return to;
//   };
// }