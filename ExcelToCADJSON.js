var jsonobj;
var ExcelToJSON = function (file) {

  this.parseExcel = function (file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = e.target.result;
      var workbook = XLSX.read(data, {
        type: 'binary'
      });
      var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[0]]);
      jsonobj = XL_row_object;

    };

    reader.onerror = function (ex) {
      console.log(ex);
    };


    reader.readAsBinaryString(file);
  };

  return this.parseExcel(file);

};

function convert() {
  let inp = document.getElementById("file");
  ExcelToJSON(inp.files[0])

  // document.getElementsByTagName("body")[0].style.cursor = "wait";
  setTimeout(() => {
    let jsonStr = JSON.stringify(jsonobj);
    jsonStr = jsonStr.replaceAll("#", "");
    // console.log(jsonStr);
    let blob = new Blob([jsonStr], { type: 'text/plain' })
    saveData(blob, "CavityCADModels.json");
    blob = null;
    // document.getElementsByTagName("body")[0].style.cursor = "auto";
  }, 1000);
}

function saveData(data, fileName) {
  var a = document.createElement("a");
  document.body.appendChild(a);
  a.style = "display: none";
  url = window.URL.createObjectURL(data);
  a.href = url;
  a.download = fileName;
  a.click();
  window.URL.revokeObjectURL(url);
  a.remove();
};