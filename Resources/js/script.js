window.onload = function () {
  loaddata($("#files").val());
};
$(document)
  .off("change", "#files")
  .on("change", "#files", function (event) {
    let fileName = $(this).val();
    loaddata(fileName);
  });

function loaddata(file) {
  var url = "./Resources/sheets/";
  url = url + file;
  var oReq = new XMLHttpRequest();
  oReq.open("GET", url, true);
  oReq.responseType = "arraybuffer";

  oReq.onload = function () {
    if (this.status == 200) {
      var arraybuffer = oReq.response;

      var data = new Uint8Array(arraybuffer);

      var arr = new Array();
      for (var i = 0; i != data.length; ++i)
        arr[i] = String.fromCharCode(data[i]);

      var bstr = arr.join("");

      var workbook = XLSX.read(bstr, {
        type: "binary",
        cellText: false,
        cellDates: true,
      });
      const workbookHeaders = XLSX.read(bstr, {
        type: "binary",
        sheetRows: 1,
      });
      var dataTable,
        htmlTable =
          '<table id="example" class="table table-striped table-bordered" width="100%"><tbody></tbody></table>';
      let optionList = document.getElementById("sheet").options;
      var sheets = workbook.SheetNames;
      $("#sheet").empty();
      const options = [];
      sheets.forEach((sheet) => {
        options.push({ text: sheet, value: sheet });
      });
      options.forEach((option) =>
        optionList.add(new Option(option.text, option.value))
      );

      function getJsonData(sheetName) {
        let ws = workbook.Sheets[sheetName];
        return XLSX.utils.sheet_to_json(ws, {
          raw: false,
          dateNF: "mm/dd/yyyy",
          defval: "",
        });
      }

      function getColumns(sheetName) {
        ("use strict");
        let columns = [];
        let ws = workbook.Sheets[sheetName];
        const columnsArray = XLSX.utils.sheet_to_json(ws, {
          header: 1,
        })[0];

        if (columnsArray) {
          columnsArray.forEach((column) => {
            columns.push({
              data: column.toString().replace(/\./g, "\\."),
              title: column,
            });
          });
        }

        return columns;
      }

      function getSheet(sheetName) {
        var minDate, maxDate;

        $.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {
          var min = minDate.val();
          var max = maxDate.val();
          var date = new Date(data[1]);

          if (
            (min === null && max === null) ||
            (min === null && date <= max) ||
            (min <= date && max === null) ||
            (min <= date && date <= max)
          ) {
            return true;
          }

          return false;
        });

        if ($.fn.DataTable.isDataTable("#example")) {
          dataTable = $("#example").DataTable();
          dataTable.destroy(true);
          $("#table-container").empty();
          $("#table-container").append(htmlTable);
        }
        minDate = new DateTime($("#min"), {
          format: "MM/DD/YYYY",
        });
        maxDate = new DateTime($("#max"), {
          format: "MM/DD/YYYY",
        });

        var data = getJsonData(sheetName);

        let columns = getColumns(sheetName);

        document.getElementById("title").innerHTML =
          "<h1>" + sheetName.toString().toUpperCase() + "</h1>";

        dataTable = $("#example").DataTable({
          bDestroy: true,
          aaData: data,
          aoColumns: columns,

          columnDefs: [
            {
              targets: "_all",
              render: function (aaData, type, row) {
                aaData = aaData + "";
                return aaData.split("\n").join("<br/>");
              },
            },
            {
              targets: 1,
              type: "date",
            },
          ],
        });

        $("#min, #max").change(function () {
          dataTable.draw();
        });
      }
      // $("#sheet").change(function () {
      //   console.log("getting sheet...");
      //   var sheet = $(this).val();
      //   getSheet(sheet);
      // });

      getSheet($("#sheet").val());

      $(document)
        .off("change", "#sheet")
        .on("change", "#sheet", function (e) {
          console.log("getting sheet...");
          var sheet = $(this).val();
          getSheet(sheet);
        });
    } else {
      console.log(this.status);
    }
  };

  oReq.send();
}
