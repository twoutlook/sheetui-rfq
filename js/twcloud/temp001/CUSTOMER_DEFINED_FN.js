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


function CUSTOM_BUTTON_CLICK_CALLBACK_FN(value, row, column, sheetId, cellObj, store) {
   // alert("DEBUG...");
    var dataVal = "CUSTOM_BUTTON_CLICK_CALLBACK_FN is called and cell button is clicked @ Row: " + row + "; Colum: " + column;
    var cells = [{sheet: sheetId, row: 12, col: 2, json: {data: dataVal}}];
//	SHEET_API.updateCells(SHEET_API_HD, cells);}
    console.log("=== XXXXX dump sheets info ===");
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
                console.log("---Key2:" + key + " Val2:" + json["cells"][key][key2]);

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
    console.log("=== CUSTOM_BUTTON_CLICK_CALLBACK_FN ===");
}
