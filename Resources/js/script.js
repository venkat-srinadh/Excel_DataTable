// var url = "https://github.com/SheetJS/test_files/blob/master/merge_cells.xls";
var url = "./Resources/sheets/file.xlsx";
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
        '<table id="example" class="display wrap" width="100%"><tbody></tbody></table>';
    let optionList = document.getElementById("sheet").options;
    var sheets = workbook.SheetNames;

    let options = [];
    sheets.forEach((sheet) => {
      options.push({ text: sheet, value: sheet });
    });
    options.forEach((option) =>
      optionList.add(new Option(option.text, option.value))
    );
    // console.log(options);

    function getJsonData(sheetName) {
      // console.log(sheetName);
      let ws = workbook.Sheets[sheetName];
      if (!ws["!merges"]) {
        console.log("merges");
      }
      ws["!merges"] = [];
      ws["!merges"].push({
        s: { c: 0, r: 0 },
        e: { c: 0, r: 7 },
      });
      // console.dir(ws, { depths: null, colors: true });
      // ws["!ref"] = "B2:Z1000";
      return XLSX.utils.sheet_to_json(ws, {
        raw: false,
        dateNF: "mm/dd/yyyy",
        blankrows: false,
      });
    }

    function getColumns(sheetName) {
      var columns = [];
      let ws = workbook.Sheets[sheetName];
      // ws["!ref"] = "B:";
      var columnsArray = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
      })[0];
      columnsArray.forEach((column) => {
        columns.push({
          data: column.toString().replace(/\./g, "\\."),
          title: column,
        });
      });

      // console.log(columns);
      return columns;
    }

    function getSheet(sheetName) {
      var minDate, maxDate;

      $.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {
        var min = minDate.val();
        var max = maxDate.val();
        var date = new Date(data[1]);
        // console.log(min);

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
      console.log(data);
      var columns = getColumns(sheetName);
      // console.log(data, columns);
      // for (var i = 0; i < columns.length; i++) {
      //   //replaces all "." with "\\." which datatables ignores
      //   columns[i].data = columns[i].data.replace(/\./g, "\\.");
      //   console.log(columns[i].data);
      // }
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
        ],
      });

      $("#min, #max").change(function () {
        dataTable.draw();
      });
    }
    $("#sheet").change(function () {
      var sheet = $(this).val();
      getSheet(sheet);
    });
    // $("#test").change(function () {
    //   d = new DateTime(document.getElementById("test"), {
    //     format: "D/M/YYYY",
    //   });
    //   console.log(d.val());
    // });

    $(document).ready(function () {
      getSheet($("#sheet").val());

      // d = new DateTime(document.getElementById("test"), {
      //   format: "D/M/YYYY",
      // });
      // console.log(d.val());
    });
  } else {
    console.log(this.status);
  }
};

oReq.send();
