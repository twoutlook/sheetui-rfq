/**
 * This is function query customer existing data ...
 */
function AUTO_FILL_CUSTOMER_DATA_BY_EMPLOYEEID(employeeId, sheetId, row, col) {

    // ok this function can call your backend and query data for employeeId and fill to here
    // In here, we just simulation it ...	
    var cells = [];
    cells.push({sheet: sheetId, row: row, col: 2, json: {data: 'John' + row + " Doe"}});
    cells.push({sheet: sheetId, row: row, col: 3, json: {render: 'dropRender', data: 'HR Dept', dropId: 1}});
    cells.push({sheet: sheetId, row: row, col: 4, json: {data: 'John' + row + '@yourCompany.com'}});
    cells.push({sheet: sheetId, row: row, col: 5, json: {data: '1 (888) 123-4567'}});
    cells.push({sheet: sheetId, row: row, col: 6, json: {data: 'Female'}});
    cells.push({sheet: sheetId, row: row, col: 7, json: {data: '1990-08-20', fm: "date", dfm: "F d, Y"}});
    cells.push({sheet: sheetId, row: row, col: 11, json: {data: 102123.34567}});
    cells.push({sheet: sheetId, row: row, col: 12, json: {data: 0.99995}});
    SHEET_API.updateCells(SHEET_API_HD, cells);
}

//http://jsfiddle.net/insin/cmewv/
//http://jsfiddle.net/h42y4ke2/21/ --- this one might work for current enviroment
function CUSTOM_BUTTON_CLICK_CALLBACK_FN002(value, row, column, sheetId, cellObj, store) {
    var dataVal = "CUSTOM_BUTTON_CLICK_CALLBACK_FN is called and cell button is clicked @ Row: " + row + "; Colum: " + column;
    var cells = [{sheet: sheetId, row: 12, col: 2, json: {data: dataVal}}];
//	SHEET_API.updateCells(SHEET_API_HD, cells);}
    console.log("=== dump sheets info ===");
//    var json = SHEET_API.getSheetTabData(SHEET_API_HD);
//    console.log(Ext.encode(json));


    //http://www.enterprisesheet.com/api/docs/manageDataAPIs/getJsonData.html
    var json = SHEET_API.getJsonData(SHEET_API_HD);
//alert(Ext.encode(json));
//    console.log(Ext.encode(json));
//    console.log("=== JSON.stringify(json) ===");
//    console.log(JSON.stringify(json))
    function visitJson(json) {
        for (var key in json) {
            console.log("Key:" + key + " Val:" + json[key]);
        }
    }

    function visitJsonCells(json) {

//        for (var key in json["cells"]) { // 畫面上會一直往右邊的col跑
//            console.log("Key:" + key + " Val:" + json["cells"][key]);
//        }

        for (var key = 0; key < 6; key++) {//
            console.log("Key:" + key + " Val:" + json["cells"][key]);
            for (var key2 in json["cells"][key]) {
                console.log("Key2:" + key + " Val2:" + json["cells"][key][key2]);

            }
        }
    }

    //  visitJsonCells(json);
    console.log("=== >>> temp001c <<< CUSTOM_BUTTON_CLICK_CALLBACK_FN ===");


    //   var htmlArr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 166, 167, 168, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 131, 106, 107, 108, 109, 110, 111, 112, 113, 114];


//----------------------
    OpenWindow = window.open("", "newwin", "height=640, width=800,toolbar=no,scrollbars=yes,menubar=no");
//写成一行
    OpenWindow.document.write(" <head>");
    OpenWindow.document.write("<TITLE>A001</TITLE>");
    OpenWindow.document.write(" <head>");
    OpenWindow.document.write('<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />');
    OpenWindow.document.write('<script  type="text/javascript" src="./js/jquery/jquery-1.12.3.min.js"></script>');
    OpenWindow.document.write('<script  type="text/javascript" src="./js/jquery/html-table-to-excel.js"></script>');

    OpenWindow.document.write(" </head>");

    OpenWindow.document.write("<BODY BGCOLOR=#ffffff>");

// OpenWindow.document.write("<h1>Hello!中文</h1>");
// OpenWindow.document.write("New window opened!");
// width: 200px;
    var strCss = '<style type="text/css"> \
        table \
        { \
            border-collapse: collapse; \
            border: none; \
            \
        } \
        td \
        { \
            border: solid #000 1px; \
        } \
    </style> \
    ';
    OpenWindow.document.write(strCss);


     OpenWindow.document.write('<h1>ver2</h1><a href="#" id="test" onClick="fnExcelReport();">download as Excel</a>');
    var str = "<table  id='myTable'>";
    var row = 0;
//   for (var i = 0; i < htmlArr.length; i++) {
//        row = htmlArr[i];
//        str += "<tr>";
//        str += '<td style="width:100px">' + row + '</td>';
//        str += '<td style="width:160px">' + strSeg[row] + '</td>';
//        str += '<td style="width:800px">' + getRowText(row) + '</td>';
//        for (var col = 0; col < submitCnt; col++) {
//            str += '<td style="width:400px">' + dataGrp[col][row] + '</td>';
//        }
//        str += "</tr>";
//    }

    var cellDataB, cellDataC, cellDataD;

    for (var i = 1; i < 115; i++) {
        cellDataA = SHEET_API.getCell(SHEET_API_HD, 1, i, 1);
        cellDataB = SHEET_API.getCell(SHEET_API_HD, 1, i, 2);
        cellDataC = SHEET_API.getCell(SHEET_API_HD, 1, i, 3);
        cellDataD = SHEET_API.getCell(SHEET_API_HD, 1, i, 4);

        str += "<tr>";
        str += "<td>" + i + "</td>";
        str += "<td>" + cellDataA.data + "</td>";
        str += "<td>" + cellDataB.data + "</td>";
        str += "<td>" + cellDataC.data + "</td>";
        str += "<td>" + cellDataD.data + "</td>";

        str += "</tr>";

    }
    str += "</table>";
//    alert(str);
    OpenWindow.document.write(str);
    OpenWindow.document.write("</BODY>");
    OpenWindow.document.write("</HTML>");
    OpenWindow.document.close();




}


function CUSTOM_BUTTON_CLICK_CALLBACK_FN(value, row, column, sheetId, cellObj, store) {
    var dataVal = "CUSTOM_BUTTON_CLICK_CALLBACK_FN is called and cell button is clicked @ Row: " + row + "; Colum: " + column;
    var cells = [{sheet: sheetId, row: 12, col: 2, json: {data: dataVal}}];
//	SHEET_API.updateCells(SHEET_API_HD, cells);}
    console.log("=== dump sheets info ===");
//    var json = SHEET_API.getSheetTabData(SHEET_API_HD);
//    console.log(Ext.encode(json));


    //http://www.enterprisesheet.com/api/docs/manageDataAPIs/getJsonData.html
    var json = SHEET_API.getJsonData(SHEET_API_HD);
//alert(Ext.encode(json));
//    console.log(Ext.encode(json));
//    console.log("=== JSON.stringify(json) ===");
//    console.log(JSON.stringify(json))
    function visitJson(json) {
        for (var key in json) {
            console.log("Key:" + key + " Val:" + json[key]);
        }
    }

    function visitJsonCells(json) {

//        for (var key in json["cells"]) { // 畫面上會一直往右邊的col跑
//            console.log("Key:" + key + " Val:" + json["cells"][key]);
//        }

        for (var key = 0; key < 6; key++) {//
            console.log("Key:" + key + " Val:" + json["cells"][key]);
            for (var key2 in json["cells"][key]) {
                console.log("Key2:" + key + " Val2:" + json["cells"][key][key2]);

            }
        }
    }



//    console.log("=== visitJson(json) === start");
//    visitJson(json);
//    console.log("=== visitJson(json) === end");
//    
//    var step2="=== visitJson(json[\"cells\"]) === ";
//    console.log(step2+"start");
//    visitJson(json["cells"]);
//    console.log(step2+" end");
//    visitJson(json);
    visitJsonCells(json);
    console.log("=== >>> temp001c <<< CUSTOM_BUTTON_CLICK_CALLBACK_FN ===");


    //   var htmlArr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 166, 167, 168, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 131, 106, 107, 108, 109, 110, 111, 112, 113, 114];


//----------------------
    OpenWindow = window.open("", "newwin", "height=640, width=800,toolbar=no,scrollbars=yes,menubar=no");
//写成一行
    OpenWindow.document.write("<TITLE>A001</TITLE>");
    OpenWindow.document.write("<BODY BGCOLOR=#ffffff>");
// OpenWindow.document.write("<h1>Hello!中文</h1>");
// OpenWindow.document.write("New window opened!");
// width: 200px;
    var strCss = '<style type="text/css"> \
        table \
        { \
            border-collapse: collapse; \
            border: none; \
            \
        } \
        td \
        { \
            border: solid #000 1px; \
        } \
    </style> \
    ';
    OpenWindow.document.write(strCss);

    var str = "<table>";
    var row = 0;
//   for (var i = 0; i < htmlArr.length; i++) {
//        row = htmlArr[i];
//        str += "<tr>";
//        str += '<td style="width:100px">' + row + '</td>';
//        str += '<td style="width:160px">' + strSeg[row] + '</td>';
//        str += '<td style="width:800px">' + getRowText(row) + '</td>';
//        for (var col = 0; col < submitCnt; col++) {
//            str += '<td style="width:400px">' + dataGrp[col][row] + '</td>';
//        }
//        str += "</tr>";
//    }

    var cellDataB, cellDataC, cellDataD;

    for (var i = 1; i < 115; i++) {
        cellDataA = SHEET_API.getCell(SHEET_API_HD, 1, i, 1);
        cellDataB = SHEET_API.getCell(SHEET_API_HD, 1, i, 2);
        cellDataC = SHEET_API.getCell(SHEET_API_HD, 1, i, 3);
        cellDataD = SHEET_API.getCell(SHEET_API_HD, 1, i, 4);

        str += "<tr>";
        str += "<td>" + i + "</td>";
        str += "<td>" + cellDataA.data + "</td>";
        str += "<td>" + cellDataB.data + "</td>";
        str += "<td>" + cellDataC.data + "</td>";
        str += "<td>" + cellDataD.data + "</td>";

        str += "</tr>";

    }
    str += "</table>";
//    alert(str);
    OpenWindow.document.write(str);
    OpenWindow.document.write("</BODY>");
    OpenWindow.document.write("</HTML>");
    OpenWindow.document.close();




}
