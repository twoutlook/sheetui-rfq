/**
 * Enterprise Spreadsheet Solutions
 * Copyright(c) FeyaSoft Inc. All right reserved.
 * info@enterpriseSheet.com
 * http://www.enterpriseSheet.com
 * 
 * Licensed under the EnterpriseSheet Commercial License.
 * http://enterprisesheet.com/license.jsp
 * 
 * You need to have a valid license key to access this file.
 */
//freezeSheet(SHEET_API_HD, 2, 2);
// var ddl_material={data: "===材质规格===", drop: Ext.encode({data: "一。铝合金,AlSi10Mg,AlSi9Cu3AlSi9Cu3,ADC12,A380,A413,A360,A356,二。锌合金,ZINC-2,ZINC-3,ZINC-5,ZINC-7ZINC-7,ZAMAK-8ZAMAK-8,三。镁合金,AZ91D,AM60,THX-AZ91D,THX-AM60"})};


 
// this is tab panel include main and details 
var centralPanel = Ext.create('enterpriseSheet.templates.CenterPanel', {
});
Ext.onReady(function () {

    Ext.QuickTips.init();

    /**
     * Define those 2 methods as global variable
     */
    SHEET_API = Ext.create('EnterpriseSheet.api.SheetAPI', {
        openFileByOnlyLoadDataFlag: true
    });

    SHEET_API_HD = SHEET_API.createSheetApp({
        withoutTitlebar: false,
        withoutSheetbar: false,
        withoutToolbar: false,
        withoutContentbar: false,
        withoutSidebar: false
    });

    //NOTE, implement freezeSheet
//    Ext.freezeSheet(SHEET_API_HD, 2, 2);
//    // this is tab panel include main and details 
//    var centralPanel = Ext.create('enterpriseSheet.templates.CenterPanel', {
//    });

    Ext.create('Ext.Viewport', {
        layout: 'border',
        items: [centralPanel],
        listeners: {
            afterlayout: function (v, layout, eOpts) {
                // westPanel.selectNode();
            }
        }
    });

    // =============================================================================================
    // ok inject data now ...
    var json = {
        fileName: "項目001",
        sheets: [{id: 1, name: "RFQ001", actived: true, color: "orange"}],
//        floatings: [
//            {sheet: 1, name: "colGroups", ftype: "colgroup", json: "[{level:1, span:[2,3]}, {level:1, span:[4,6]}]"},
//        ],
        cells: [
            {sheet: 1, row: 0, col: 0, json: {height: 20, va: "middle"}},
            {sheet: 1, row: 0, col: 1, json: {data: "A", width: 50}},
            {sheet: 1, row: 0, col: 2, json: {data: "B", width: 300}},
            {sheet: 1, row: 0, col: 3, json: {data: "C", width: 150}},
            {sheet: 1, row: 0, col: 4, json: {data: "D", width: 150}},
            {sheet: 1, row: 0, col: 5, json: {data: "E", width: 150}},
            {sheet: 1, row: 0, col: 6, json: {data: "F", width: 150}},
            {sheet: 1, row: 0, col: 7, json: {data: "G", width: 150}},
            {sheet: 1, row: 1, col: 1, json: {data: ""}},
            {sheet: 1, row: 1, col: 2, json: {data: "實現下拉"}},
            {sheet: 1, row: 1, col: 3, json: {data: "0"}},
            {sheet: 1, row: 2, col: 1, json: {data: ""}},
            {sheet: 1, row: 2, col: 2, json: {data: ""}},
            {sheet: 1, row: 2, col: 3, json: {data: "0"}},
            {sheet: 1, row: 3, col: 1, json: {data: ""}},
            {sheet: 1, row: 3, col: 2, json: {data: "接受询价日期："}},
            {sheet: 1, row: 3, col: 3, json: {data: "0"}},
            {sheet: 1, row: 4, col: 1, json: {data: ""}},
            {sheet: 1, row: 4, col: 2, json: {data: "业务担当："}},
            {sheet: 1, row: 4, col: 3, json: {data: "0"}},
            {sheet: 1, row: 5, col: 1, json: {data: ""}},
            {sheet: 1, row: 5, col: 2, json: {data: ""}},
            {sheet: 1, row: 5, col: 3, json: {data: "0"}},
            {sheet: 1, row: 6, col: 1, json: {data: ""}},
            {sheet: 1, row: 6, col: 2, json: {data: "产品简图 "}},
            {sheet: 1, row: 6, col: 3, json: {data: "0"}},
            {sheet: 1, row: 7, col: 1, json: {data: ""}},
            {sheet: 1, row: 7, col: 2, json: {data: "项目 料号 "}},
            {sheet: 1, row: 7, col: 3, json: {data: "0"}},
            {sheet: 1, row: 8, col: 1, json: {data: ""}},
            {sheet: 1, row: 8, col: 2, json: {data: "品名 / 图纸版本号："}},
            {sheet: 1, row: 8, col: 3, json: {data: "0"}},
            {sheet: 1, row: 9, col: 1, json: {data: ""}},
            {sheet: 1, row: 9, col: 2, json: {data: "产品外形尺寸"}},
            {sheet: 1, row: 9, col: 3, json: {data: "0"}},
            {sheet: 1, row: 10, col: 1, json: {data: ""}},
            {sheet: 1, row: 10, col: 2, json: {data: "材质规格： "}},
            {sheet: 1, row: 10, col: 3, json: {data: "???==="}},
            {sheet: 1, row: 11, col: 1, json: {data: ""}},
            {sheet: 1, row: 11, col: 2, json: {data: "产品单重（g）："}},
            {sheet: 1, row: 11, col: 3, json: {data: "0"}},
            {sheet: 1, row: 12, col: 1, json: {data: ""}},
            {sheet: 1, row: 12, col: 2, json: {data: "表面要求： "}},
            {sheet: 1, row: 12, col: 3, json: {data: "==="}},
            {sheet: 1, row: 13, col: 1, json: {data: ""}},
            {sheet: 1, row: 13, col: 2, json: {data: "年需求量："}},
            {sheet: 1, row: 13, col: 3, json: {data: "12000"}},
            {sheet: 1, row: 14, col: 1, json: {data: ""}},
            {sheet: 1, row: 14, col: 2, json: {data: ""}},
            {sheet: 1, row: 14, col: 3, json: {data: "0"}},
            {sheet: 1, row: 15, col: 1, json: {data: ""}},
            {sheet: 1, row: 15, col: 2, json: {data: "模具费用： "}},
            {sheet: 1, row: 15, col: 3, json: {data: "0"}},
            {sheet: 1, row: 16, col: 1, json: {data: ""}},
            {sheet: 1, row: 16, col: 2, json: {data: "模穴数 "}},
            {sheet: 1, row: 16, col: 3, json: {data: "1"}},
            {sheet: 1, row: 17, col: 1, json: {data: ""}},
            {sheet: 1, row: 17, col: 2, json: {data: "侧滑模数/油压抽芯数 "}},
            {sheet: 1, row: 17, col: 3, json: {data: "0"}},
            {sheet: 1, row: 18, col: 1, json: {data: ""}},
            {sheet: 1, row: 18, col: 2, json: {data: "模具寿命："}},
            {sheet: 1, row: 18, col: 3, json: {data: "0"}},
            {sheet: 1, row: 19, col: 1, json: {data: ""}},
            {sheet: 1, row: 19, col: 2, json: {data: "压铸模费用： "}},
            {sheet: 1, row: 19, col: 3, json: {data: "0"}},
            {sheet: 1, row: 20, col: 1, json: {data: ""}},
            {sheet: 1, row: 20, col: 2, json: {data: "切边模费用： "}},
            {sheet: 1, row: 20, col: 3, json: {data: "0"}},
            {sheet: 1, row: 21, col: 1, json: {data: ""}},
            {sheet: 1, row: 21, col: 2, json: {data: "加工夹治具费用： "}},
            {sheet: 1, row: 21, col: 3, json: {data: "0"}},
            {sheet: 1, row: 22, col: 1, json: {data: ""}},
            {sheet: 1, row: 22, col: 2, json: {data: "烤漆夹治具费用： "}},
            {sheet: 1, row: 22, col: 3, json: {data: "0"}},
            {sheet: 1, row: 23, col: 1, json: {data: ""}},
            {sheet: 1, row: 23, col: 2, json: {data: "模具总价（人民币）： "}},
            {sheet: 1, row: 23, col: 3, json: {data: "0"}},
            {sheet: 1, row: 24, col: 1, json: {data: ""}},
            {sheet: 1, row: 24, col: 2, json: {data: "模具总价（USD）： "}},
            {sheet: 1, row: 24, col: 3, json: {data: "0"}},
            {sheet: 1, row: 25, col: 1, json: {data: ""}},
            {sheet: 1, row: 25, col: 2, json: {data: ""}},
            {sheet: 1, row: 25, col: 3, json: {data: "0"}},
            {sheet: 1, row: 26, col: 1, json: {data: ""}},
            {sheet: 1, row: 26, col: 2, json: {data: ""}},
            {sheet: 1, row: 26, col: 3, json: {data: "0"}},
            {sheet: 1, row: 27, col: 1, json: {data: ""}},
            {sheet: 1, row: 27, col: 2, json: {data: "产品单价： "}},
            {sheet: 1, row: 27, col: 3, json: {data: "0"}},
            {sheet: 1, row: 28, col: 1, json: {data: "* "}},
            {sheet: 1, row: 28, col: 2, json: {data: "1)材料费 ： "}},
            {sheet: 1, row: 28, col: 3, json: {data: "0"}},
            {sheet: 1, row: 29, col: 1, json: {data: "1-1）"}},
            {sheet: 1, row: 29, col: 2, json: {data: "材质规格："}},
            {sheet: 1, row: 29, col: 3, json: {data: "==="}},
            {sheet: 1, row: 30, col: 1, json: {data: "1-2）"}},
            {sheet: 1, row: 30, col: 2, json: {data: "材料单价/KG :"}},
            {sheet: 1, row: 30, col: 3, json: {data: "0"}},
            {sheet: 1, row: 31, col: 1, json: {data: "1-3）"}},
            {sheet: 1, row: 31, col: 2, json: {data: "单重（g）："}},
            {sheet: 1, row: 31, col: 3, json: {data: "0"}},
            {sheet: 1, row: 32, col: 1, json: {data: "1-4）"}},
            {sheet: 1, row: 32, col: 2, json: {data: "产品材料费（净重价格）： "}},
            {sheet: 1, row: 32, col: 3, json: {data: "0"}},
            {sheet: 1, row: 33, col: 1, json: {data: "1-5）"}},
            {sheet: 1, row: 33, col: 2, json: {data: "料柄流道渣包等废料重量(g)： "}},
            {sheet: 1, row: 33, col: 3, json: {data: "0"}},
            {sheet: 1, row: 34, col: 1, json: {data: "1-6）"}},
            {sheet: 1, row: 34, col: 2, json: {data: "料柄流道渣包比率 ： "}},
            {sheet: 1, row: 34, col: 3, json: {data: "0"}},
            {sheet: 1, row: 35, col: 1, json: {data: "1-7）"}},
            {sheet: 1, row: 35, col: 2, json: {data: "料柄流道渣包等废料价格/Kg： "}},
            {sheet: 1, row: 35, col: 3, json: {data: "0"}},
            {sheet: 1, row: 36, col: 1, json: {data: "1-8）"}},
            {sheet: 1, row: 36, col: 2, json: {data: "料柄流道渣包回收价差损失额 ： "}},
            {sheet: 1, row: 36, col: 3, json: {data: "0"}},
            {sheet: 1, row: 37, col: 1, json: {data: "1-9）"}},
            {sheet: 1, row: 37, col: 2, json: {data: "压铸熔炼材料氧化损耗(量）2% "}},
            {sheet: 1, row: 37, col: 3, json: {data: "0"}},
            {sheet: 1, row: 38, col: 1, json: {data: ""}},
            {sheet: 1, row: 38, col: 2, json: {data: "材料费　小计： "}},
            {sheet: 1, row: 38, col: 3, json: {data: "0"}},
            {sheet: 1, row: 39, col: 1, json: {data: "* "}},
            {sheet: 1, row: 39, col: 2, json: {data: "2)压铸费(含切边）： "}},
            {sheet: 1, row: 39, col: 3, json: {data: "0"}},
            {sheet: 1, row: 40, col: 1, json: {data: "2-1）"}},
            {sheet: 1, row: 40, col: 2, json: {data: "适用机型： "}},
            {sheet: 1, row: 40, col: 3, json: {data: "==="}},
            {sheet: 1, row: 41, col: 1, json: {data: "2-2）"}},
            {sheet: 1, row: 41, col: 2, json: {data: "设备费率/H： "}},
            {sheet: 1, row: 41, col: 3, json: {data: "0"}},
            {sheet: 1, row: 42, col: 1, json: {data: "2-3）"}},
            {sheet: 1, row: 42, col: 2, json: {data: "工时（秒）/只 ："}},
            {sheet: 1, row: 42, col: 3, json: {data: "0"}},
            {sheet: 1, row: 43, col: 1, json: {data: "2-4）"}},
            {sheet: 1, row: 43, col: 2, json: {data: "产能 只/H ："}},
            {sheet: 1, row: 43, col: 3, json: {data: "0"}},
            {sheet: 1, row: 44, col: 1, json: {data: "2-5）"}},
            {sheet: 1, row: 44, col: 2, json: {data: "设备费/只： "}},
            {sheet: 1, row: 44, col: 3, json: {data: "0"}},
            {sheet: 1, row: 45, col: 1, json: {data: "2-6）"}},
            {sheet: 1, row: 45, col: 2, json: {data: "人工费/只： "}},
            {sheet: 1, row: 45, col: 3, json: {data: "0"}},
            {sheet: 1, row: 46, col: 1, json: {data: "2-7）"}},
            {sheet: 1, row: 46, col: 2, json: {data: "熔炼费（能耗）： "}},
            {sheet: 1, row: 46, col: 3, json: {data: "0"}},
            {sheet: 1, row: 47, col: 1, json: {data: "2-8）"}},
            {sheet: 1, row: 47, col: 2, json: {data: "良品率： "}},
            {sheet: 1, row: 47, col: 3, json: {data: "0.95"}},
            {sheet: 1, row: 48, col: 1, json: {data: ""}},
            {sheet: 1, row: 48, col: 2, json: {data: "压铸费　小计： "}},
            {sheet: 1, row: 48, col: 3, json: {data: "0"}},
            {sheet: 1, row: 49, col: 1, json: {data: "* "}},
            {sheet: 1, row: 49, col: 2, json: {data: "3)毛刺处理费（含粗磨）"}},
            {sheet: 1, row: 49, col: 3, json: {data: "0"}},
            {sheet: 1, row: 50, col: 1, json: {data: "3-1）"}},
            {sheet: 1, row: 50, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 50, col: 3, json: {data: "0"}},
            {sheet: 1, row: 51, col: 1, json: {data: "3-2）"}},
            {sheet: 1, row: 51, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 51, col: 3, json: {data: "35"}},
            {sheet: 1, row: 52, col: 1, json: {data: ""}},
            {sheet: 1, row: 52, col: 2, json: {data: "毛刺处理费　小计： "}},
            {sheet: 1, row: 52, col: 3, json: {data: "0"}},
            {sheet: 1, row: 53, col: 1, json: {data: "* "}},
            {sheet: 1, row: 53, col: 2, json: {data: "4)外观打磨费（入/溢料口, 分模线及一般外观瑕疵等部位做打磨）"}},
            {sheet: 1, row: 53, col: 3, json: {data: "0"}},
            {sheet: 1, row: 54, col: 1, json: {data: "4-1）"}},
            {sheet: 1, row: 54, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 54, col: 3, json: {data: "0"}},
            {sheet: 1, row: 55, col: 1, json: {data: "4-2）"}},
            {sheet: 1, row: 55, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 55, col: 3, json: {data: "45"}},
            {sheet: 1, row: 56, col: 1, json: {data: "4-3）"}},
            {sheet: 1, row: 56, col: 2, json: {data: "打磨费： "}},
            {sheet: 1, row: 56, col: 3, json: {data: "0"}},
            {sheet: 1, row: 57, col: 1, json: {data: "4-4）"}},
            {sheet: 1, row: 57, col: 2, json: {data: "良品率： "}},
            {sheet: 1, row: 57, col: 3, json: {data: "0.95"}},
            {sheet: 1, row: 58, col: 1, json: {data: ""}},
            {sheet: 1, row: 58, col: 2, json: {data: ""}},
            {sheet: 1, row: 58, col: 3, json: {data: "0"}},
            {sheet: 1, row: 59, col: 1, json: {data: ""}},
            {sheet: 1, row: 59, col: 2, json: {data: "外观打磨费　小计："}},
            {sheet: 1, row: 59, col: 3, json: {data: "0"}},
            {sheet: 1, row: 60, col: 1, json: {data: "* "}},
            {sheet: 1, row: 60, col: 2, json: {data: "5)平整度整形费 "}},
            {sheet: 1, row: 60, col: 3, json: {data: "0"}},
            {sheet: 1, row: 61, col: 1, json: {data: "5-1）"}},
            {sheet: 1, row: 61, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 61, col: 3, json: {data: "0"}},
            {sheet: 1, row: 62, col: 1, json: {data: "5-2）"}},
            {sheet: 1, row: 62, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 62, col: 3, json: {data: "45"}},
            {sheet: 1, row: 63, col: 1, json: {data: "5-3）"}},
            {sheet: 1, row: 63, col: 2, json: {data: "人工费/只 "}},
            {sheet: 1, row: 63, col: 3, json: {data: "0"}},
            {sheet: 1, row: 64, col: 1, json: {data: ""}},
            {sheet: 1, row: 64, col: 2, json: {data: "平整度整形费　小计： "}},
            {sheet: 1, row: 64, col: 3, json: {data: "0"}},
            {sheet: 1, row: 65, col: 1, json: {data: "* "}},
            {sheet: 1, row: 65, col: 2, json: {data: "6)机加工 "}},
            {sheet: 1, row: 65, col: 3, json: {data: "==="}},
            {sheet: 1, row: 66, col: 1, json: {data: "6-1）"}},
            {sheet: 1, row: 66, col: 2, json: {data: "傳統机加工工时（秒）/只 ： "}},
            {sheet: 1, row: 66, col: 3, json: {data: "0"}},
            {sheet: 1, row: 67, col: 1, json: {data: "6-2）"}},
            {sheet: 1, row: 67, col: 2, json: {data: "傳統机加工费率/H： "}},
            {sheet: 1, row: 67, col: 3, json: {data: "0"}},
            {sheet: 1, row: 68, col: 1, json: {data: "6-3）"}},
            {sheet: 1, row: 68, col: 2, json: {data: "傳統机加工良品率： "}},
            {sheet: 1, row: 68, col: 3, json: {data: "0.95"}},
            {sheet: 1, row: 69, col: 1, json: {data: ""}},
            {sheet: 1, row: 69, col: 2, json: {data: "机加工　小计： "}},
            {sheet: 1, row: 69, col: 3, json: {data: "0"}},
            {sheet: 1, row: 70, col: 1, json: {data: "* "}},
            {sheet: 1, row: 70, col: 2, json: {data: "7)冷喷.热烧毛刺 &费率/分： "}},
            {sheet: 1, row: 70, col: 3, json: {data: "0"}},
            {sheet: 1, row: 71, col: 1, json: {data: "7-1）"}},
            {sheet: 1, row: 71, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 71, col: 3, json: {data: "0"}},
            {sheet: 1, row: 72, col: 1, json: {data: "7-2）"}},
            {sheet: 1, row: 72, col: 2, json: {data: "加工费率/H： "}},
            {sheet: 1, row: 72, col: 3, json: {data: "35"}},
            {sheet: 1, row: 73, col: 1, json: {data: ""}},
            {sheet: 1, row: 73, col: 2, json: {data: "冷喷热烧毛刺　小计："}},
            {sheet: 1, row: 73, col: 3, json: {data: "0"}},
            {sheet: 1, row: 74, col: 1, json: {data: "* "}},
            {sheet: 1, row: 74, col: 2, json: {data: "8)喷沙.抛丸.震动研磨 &费率 "}},
            {sheet: 1, row: 74, col: 3, json: {data: "==="}},
            {sheet: 1, row: 75, col: 1, json: {data: "8-1）"}},
            {sheet: 1, row: 75, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 75, col: 3, json: {data: "0"}},
            {sheet: 1, row: 76, col: 1, json: {data: "8-2）"}},
            {sheet: 1, row: 76, col: 2, json: {data: "加工费率/H： "}},
            {sheet: 1, row: 76, col: 3, json: {data: "0"}},
            {sheet: 1, row: 77, col: 1, json: {data: ""}},
            {sheet: 1, row: 77, col: 2, json: {data: "喷沙抛丸震动研磨　小计： "}},
            {sheet: 1, row: 77, col: 3, json: {data: "0"}},
            {sheet: 1, row: 78, col: 1, json: {data: "* "}},
            {sheet: 1, row: 78, col: 2, json: {data: "9)皮膜处理 &费率 "}},
            {sheet: 1, row: 78, col: 3, json: {data: "0"}},
            {sheet: 1, row: 79, col: 1, json: {data: "9-1）"}},
            {sheet: 1, row: 79, col: 2, json: {data: "皮膜处理： "}},
            {sheet: 1, row: 79, col: 3, json: {data: "==="}},
            {sheet: 1, row: 80, col: 1, json: {data: "9-2）"}},
            {sheet: 1, row: 80, col: 2, json: {data: "工件表面面积 (算单面）DM2： "}},
            {sheet: 1, row: 80, col: 3, json: {data: "0"}},
            {sheet: 1, row: 81, col: 1, json: {data: "9-3）"}},
            {sheet: 1, row: 81, col: 2, json: {data: "加工费率/H： "}},
            {sheet: 1, row: 81, col: 3, json: {data: "0"}},
            {sheet: 1, row: 82, col: 1, json: {data: ""}},
            {sheet: 1, row: 82, col: 2, json: {data: "皮膜处理　小计： "}},
            {sheet: 1, row: 82, col: 3, json: {data: "0"}},
            {sheet: 1, row: 83, col: 1, json: {data: ""}},
            {sheet: 1, row: 83, col: 2, json: {data: ""}},
            {sheet: 1, row: 83, col: 3, json: {data: "0"}},
            {sheet: 1, row: 84, col: 1, json: {data: "* "}},
            {sheet: 1, row: 84, col: 2, json: {data: "10)粉体烤漆.液体烤漆.丝网印： "}},
            {sheet: 1, row: 84, col: 3, json: {data: "0"}},
            {sheet: 1, row: 85, col: 1, json: {data: "10-1）"}},
            {sheet: 1, row: 85, col: 2, json: {data: "表面要求"}},
            {sheet: 1, row: 85, col: 3, json: {data: "==="}},
            {sheet: 1, row: 86, col: 1, json: {data: "10-2）"}},
            {sheet: 1, row: 86, col: 2, json: {data: "工件表面积(算需喷漆面积）DM2： "}},
            {sheet: 1, row: 86, col: 3, json: {data: "0"}},
            {sheet: 1, row: 87, col: 1, json: {data: "10-3）"}},
            {sheet: 1, row: 87, col: 2, json: {data: "漆材料费/dm： "}},
            {sheet: 1, row: 87, col: 3, json: {data: "0"}},
            {sheet: 1, row: 88, col: 1, json: {data: "10-4）"}},
            {sheet: 1, row: 88, col: 2, json: {data: "烤漆设备+人工費率/dm： "}},
            {sheet: 1, row: 88, col: 3, json: {data: "0"}},
            {sheet: 1, row: 89, col: 1, json: {data: "10-5）"}},
            {sheet: 1, row: 89, col: 2, json: {data: "良品率： "}},
            {sheet: 1, row: 89, col: 3, json: {data: "0.95"}},
            {sheet: 1, row: 90, col: 1, json: {data: "10-6）"}},
            {sheet: 1, row: 90, col: 2, json: {data: "产品难度系数 K 值：K = 0.8-1.5 "}},
            {sheet: 1, row: 90, col: 3, json: {data: "1"}},
            {sheet: 1, row: 91, col: 1, json: {data: ""}},
            {sheet: 1, row: 91, col: 2, json: {data: "粉体液体烤漆丝网印　小计： "}},
            {sheet: 1, row: 91, col: 3, json: {data: "0"}},
            {sheet: 1, row: 92, col: 1, json: {data: "* "}},
            {sheet: 1, row: 92, col: 2, json: {data: "11)其它特殊处理： "}},
            {sheet: 1, row: 92, col: 3, json: {data: "0"}},
            {sheet: 1, row: 93, col: 1, json: {data: "11-1）"}},
            {sheet: 1, row: 93, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 93, col: 3, json: {data: "0"}},
            {sheet: 1, row: 94, col: 1, json: {data: "11-2）"}},
            {sheet: 1, row: 94, col: 2, json: {data: "费用/只 "}},
            {sheet: 1, row: 94, col: 3, json: {data: "0"}},
            {sheet: 1, row: 95, col: 1, json: {data: ""}},
            {sheet: 1, row: 95, col: 2, json: {data: "其它特殊处理　小计： "}},
            {sheet: 1, row: 95, col: 3, json: {data: "0"}},
            {sheet: 1, row: 96, col: 1, json: {data: "* "}},
            {sheet: 1, row: 96, col: 2, json: {data: "12)气密性测试 ： "}},
            {sheet: 1, row: 96, col: 3, json: {data: "0"}},
            {sheet: 1, row: 97, col: 1, json: {data: "12-1）"}},
            {sheet: 1, row: 97, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 97, col: 3, json: {data: "0"}},
            {sheet: 1, row: 98, col: 1, json: {data: "12-2）"}},
            {sheet: 1, row: 98, col: 2, json: {data: "气测费/只 "}},
            {sheet: 1, row: 98, col: 3, json: {data: "0"}},
            {sheet: 1, row: 99, col: 1, json: {data: ""}},
            {sheet: 1, row: 99, col: 2, json: {data: "气密性测试　小计： "}},
            {sheet: 1, row: 99, col: 3, json: {data: "0"}},
            {sheet: 1, row: 100, col: 1, json: {data: "* "}},
            {sheet: 1, row: 100, col: 2, json: {data: "13)筛选包装 & 费率 /H： "}},
            {sheet: 1, row: 100, col: 3, json: {data: "35"}},
            {sheet: 1, row: 101, col: 1, json: {data: "13-1）"}},
            {sheet: 1, row: 101, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 101, col: 3, json: {data: "0"}},
            {sheet: 1, row: 102, col: 1, json: {data: "13-2）"}},
            {sheet: 1, row: 102, col: 2, json: {data: "人工费/只 "}},
            {sheet: 1, row: 102, col: 3, json: {data: "0"}},
            {sheet: 1, row: 103, col: 1, json: {data: "13-3）"}},
            {sheet: 1, row: 103, col: 2, json: {data: "包材费： "}},
            {sheet: 1, row: 103, col: 3, json: {data: "0"}},
            {sheet: 1, row: 104, col: 1, json: {data: ""}},
            {sheet: 1, row: 104, col: 2, json: {data: "筛选包装　小计："}},
            {sheet: 1, row: 104, col: 3, json: {data: "0"}},
            {sheet: 1, row: 105, col: 1, json: {data: ""}},
            {sheet: 1, row: 105, col: 2, json: {data: "各工序费用　合计："}},
            {sheet: 1, row: 105, col: 3, json: {data: "0"}},
            {sheet: 1, row: 106, col: 1, json: {data: "* "}},
            {sheet: 1, row: 106, col: 2, json: {data: "管销利润"}},
            {sheet: 1, row: 106, col: 3, json: {data: "0"}},
            {sheet: 1, row: 107, col: 1, json: {data: "* "}},
            {sheet: 1, row: 107, col: 2, json: {data: "其他五金配件： "}},
            {sheet: 1, row: 107, col: 3, json: {data: "0"}},
            {sheet: 1, row: 108, col: 1, json: {data: "* "}},
            {sheet: 1, row: 108, col: 2, json: {data: "组装费： "}},
            {sheet: 1, row: 108, col: 3, json: {data: "0"}},
            {sheet: 1, row: 109, col: 1, json: {data: ""}},
            {sheet: 1, row: 109, col: 2, json: {data: "组装 小计： "}},
            {sheet: 1, row: 109, col: 3, json: {data: "0"}},
            {sheet: 1, row: 110, col: 1, json: {data: ""}},
            {sheet: 1, row: 110, col: 2, json: {data: "合計（未稅）："}},
            {sheet: 1, row: 110, col: 3, json: {data: "0"}},
            {sheet: 1, row: 111, col: 1, json: {data: ""}},
            {sheet: 1, row: 111, col: 2, json: {data: "VAT 税金：17% "}},
            {sheet: 1, row: 111, col: 3, json: {data: "0"}},
            {sheet: 1, row: 112, col: 1, json: {data: ""}},
            {sheet: 1, row: 112, col: 2, json: {data: "海外运输费用： "}},
            {sheet: 1, row: 112, col: 3, json: {data: "0"}},
            {sheet: 1, row: 113, col: 1, json: {data: ""}},
            {sheet: 1, row: 113, col: 2, json: {data: "总计（含税）： "}},
            {sheet: 1, row: 113, col: 3, json: {data: "0"}},
        ]
    };

    SHEET_API.loadData(SHEET_API_HD, json, null, this);
    SHEET_API.setFocus(SHEET_API_HD, 2, 1);
//
//    var json123 = SHEET_API.getJsonData(SHEET_API_HD);
//    alert(Ext.encode(json123));



    // add event listener - this shows the code to add customer function 
    var sheet = SHEET_API_HD.sheet;
    var editor = sheet.getEditor();
    editor.on('quit', function (editor, sheetId, row, col) {
        if (col === 1) {
            // this is the method to query customer existing backend and auto fill data
            var employeeId = SHEET_API.getCellValue(SHEET_API_HD, sheetId, row, col).data;
            if (employeeId)
                AUTO_FILL_CUSTOMER_DATA_BY_EMPLOYEEID(employeeId, sheetId, row, col);
        }
    }, this);

    // add cell on select event ...
    /**
     var sm = sheet.getSelectionModel();
     sm.on('selectionchange', function(startPos, endPos, region, sm) {
     if (startPos.row == endPos.row && startPos.col == endPos.col && startPos.col == 8) {
     this.customEditor = Ext.create('customer.CellEditor', {
     sheetId: region.sheetId,
     row: startPos.row,
     col: startPos.col
     });
     this.customEditor.popup();
     }
     }, this);
     **/

    sheet.on('_beforeeditcell_', function (sheetId, row, col, cellData, sheet) {

//        var json123 = SHEET_API.getJsonData(SHEET_API_HD);
////    alert(Ext.encode(json123));
//        console.log(Ext.encode(json123));


        var contactWin = Ext.create('customer.CellEditor', {
            sheetId: sheetId,
            row: row,
            col: col,
            cellData: cellData,
            sheet: sheet
        });
        contactWin.popup();
        return false;
    }, this);

});
	