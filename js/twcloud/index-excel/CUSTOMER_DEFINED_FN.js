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
    //NOTE: 5/19 11:54 end user asking to export as real Excel with functions on cells
    $.post("php-excel/generate-excel.php", {
        name: "Mark",
        city: "昆山"
    },
            function (data, status) {
//                alert("Data: " + data + "\nStatus: " + status);
                if (status === "success") {
                    console.log(data);

                    var cells = [];
                    cells.push({
                        sheet: 1,
                        row: 5,
                        col: 1,
                        json: { data: "下載", link:data} 
//                        json: { data:data} 
                    });
                    SHEET_API.updateCells(SHEET_API_HD, cells);


//                    alert(data);
                } else {
                    console.log("CUSTOM_BUTTON_CLICK_CALLBACK_FN002 failed");
//                  
//                    alert("not successful");
                }

//             $("#div-msg").html(data).css("color", "gray").css("fontSize", "16px").show(100);
            });

    return;

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
//     OpenWindow.document.write('<h1>ver2</h1><a href="#" id="test" onClick="fnExcelReport();">download as Excel</a>');
    OpenWindow.document.write('<h1>ver3 <a href="#" id="test" onClick="fnExcelReport();">下載 Excel 檔案</a></h1>');
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

    function isItemInArr(val, arr) {
        // var arrA=[0,1,1,1,2,2,2];
        for (var i = 0; i < arr.length; i++) {
            if (val === arr[i]) {
                return true;
            }
        }
        return false;


    }
    function getSegmentFormat(i, data) {
        var arrA = [4, 113];
        if (isItemInArr(i, arrA)) {
            return  "<td></td>"
        }
        return  "<td align='center'>" + data + "</td>";

    }
//    var colorStep = "#A9BCF5";
//    var colorStepEnd = "#E6E6E6";
//    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
//    var colorDdl = "#F9E79F"; //#82E0AA  
//    var colorInput = "#F4D03F"; // 

    var arrStepEnd = [23, 24, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111, 112];


    function getTitleFormat(i, data) {
        var arrA = [15, 28, 39, 49, 53, 60, 65, 70, 74, 78, 84, 92, 96, 100];
        if (isItemInArr(i, arrA)) {
            return  "<td align='center' style='background:#A9BCF5'>" + data + "</td>";
        } else if (isItemInArr(i, arrStepEnd)) {
            return  "<td align='center' style='background:#E6E6E6'>" + data + "</td>";
        }
        return  "<td align='center'>" + data + "</td>";

    }
    function getDataAlign(i) {
        var arrDataCenter = [7, 8, 9, 10, 12, 17, 18, 29, 40, 65, 70, 74, 79, 85, ];
        if (isItemInArr(i, arrDataCenter)) {
            return  " align='center' ";//#F9E79F
        }
        return " align='right' ";
    }
    function getDataFormat(i, data) {
        var arrDdl = [10, 12, 40, 65, 70, 74, 79];
        var arrInput = [11, 16, 19, 20, 21, 22, 33, 42, 47, 50, 54, 57, 61, 68, 71, 75, 80, 81, 82, 86, 87, 88, 89, 90, 93, 94, 97, 101, 103, 108, 109];

        var strAlign = getDataAlign(i);


        //      var colorStepEnd = "#E6E6E6";
        if (isItemInArr(i, arrInput)) {
            return  "<td " + strAlign + " style='background:#F4D03F'>" + data + "</td>";//#F9E79F
        } else if (isItemInArr(i, arrDdl)) {
            return  "<td " + strAlign + "  style='background:#F9E79F'>" + data + "</td>";//#F9E79F
        } else if (isItemInArr(i, arrStepEnd)) {
            return  "<td " + strAlign + "  style='background:#E6E6E6'>" + data + "</td>";//#F9E79F
        }
        return  "<td " + strAlign + ">" + data + "</td>";

    }



    var cellDataB, cellDataC, cellDataD;
    var cellDataX = new Array(6);
    for (var i = 1; i < 113; i++) {
        cellDataA = SHEET_API.getCell(SHEET_API_HD, 1, i, 1);
        cellDataB = SHEET_API.getCell(SHEET_API_HD, 1, i, 2);
//        cellDataC = SHEET_API.getCell(SHEET_API_HD, 1, i, 3);
//        cellDataD = SHEET_API.getCell(SHEET_API_HD, 1, i, 4);
        for (var arrIndex = 0; arrIndex < 6; arrIndex++) {
            cellDataX[arrIndex] = SHEET_API.getCell(SHEET_API_HD, 1, i, 3 + arrIndex);
        }


        str += "<tr>";
//        str += "<td>" + i + "</td>";

        str += getSegmentFormat(i, cellDataA.data);//不顯示導出字樣

//        str += "<td>" + cellDataB.data + "</td>";
        str += getTitleFormat(i, cellDataB.data);//不顯示導出字樣




//        str += "<td>" + cellDataC.data + "</td>";
//        str += "<td>" + cellDataD.data + "</td>";
        for (var arrIndex = 0; arrIndex < 6; arrIndex++) {

//             str += "<td>" + cellDataX[arrIndex].data + "</td>";
            str += getDataFormat(i, cellDataX[arrIndex].data);
        }

        str += "</tr>";
    }
    str += "</table>";
//    alert(str);
    OpenWindow.document.write(str);
    //fnExcelReport();
//     OpenWindow.document.write("<script>alert('Can we?');fnExcelReport();alert('Did it work?')</script>");
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
