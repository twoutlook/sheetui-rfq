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
    // this is tab panel include main and details 
    var centralPanel = Ext.create('enterpriseSheet.templates.CenterPanel', {
    });
    Ext.create('Ext.Viewport', {
        layout: 'border',
        items: [centralPanel],
        listeners: {
            afterlayout: function (v, layout, eOpts) {
                // westPanel.selectNode();
            }
        }
    });
    //  NOTE:
    //  
    var colorStep = "#A9BCF5";
    var colorStepEnd = "#E6E6E6";
    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
    var colorDdl = "#F9E79F"; //#82E0AA  
    var colorInput = "#F4D03F"; // 
    //
    //
    //
//        var ddlMaterial = {data: "xxx", drop: Ext.encode({data: "xxx,yyy,zzz"})};
    var ddlMaterial = {bgc: colorDdl, ta: "center", data: "===材质规格===", drop: Ext.encode({data: "一。铝合金,AlSi10Mg,AlSi9Cu3,ADC12,A380,A413,A360,A356,二。锌合金 ,ZINC-2,ZINC-3,ZINC-5,ZINC-7,ZAMAK-8,三。镁合金 ,AZ91D,AM60,THX-AZ91D,THX-AM60"})};
    var ddlSurface = {bgc: colorDdl, ta: "center", data: "===表面要求===", drop: Ext.encode({data: "单清洗,烤漆前皮膜（含清洗）,铝合金一般皮膜（48H）,锌合金一般皮膜（48H),镁合金一般皮膜（24H),铝合金-特殊皮膜（   H）,锌合金-特殊皮膜（   H) ,特殊导电皮膜(  欧姆) ,粉體烤漆-A+級,粉體烤漆-A級 ,粉體烤漆-B級 ,液體烤漆-A+級 ,液體烤漆-A級 ,液體烤漆-B級 ,阳极氧化-A级 , 阳极氧化-B级 ,电泳-A级,电泳-B级 ,掛鍍-A級,掛鍍-B 級,滾鍍-A級,滾鍍-B級 ,高清洁度清洗（600um）,高清洁度清洗（400um）,高清洁度清洗（200um）,清洗鉻酸"})};
    var ddlMachine = {bgc: colorDdl, ta: "center", data: "===适用机型===", drop: Ext.encode({data: "一。铝合金压铸合模费：,铝-125T,鋁-150T/160T/200T,鋁-250T/280T,鋁-350T/340T/400T,鋁-550T/530T/500T,鋁-630T/650T,鋁-800T/900T,鋁-1250T,鋁-1600T,鋁-2000T,鋁-2500T,鋁-3000T,二。锌合金压铸合模费：,鋅-快速机/4轴机,鋅-10T/5T,鋅-15T/20T,鋅-25T/30T/40T,鋅-50T/60T,鋅-80T/100T/125,鋅-150T,鋅-185/200T,鋅-250T,鋅-300T,三。镁合金压铸合模费：,鎂-150T,鎂-340T/400T,鎂-530T/660T,THX-鎂280T,THX-鎂650T"})};
//DOING...喷沙sand blasting 抛丸.震动研磨
    var ddlSand = {bgc: colorDdl, ta: "center", data: "===喷砂,抛丸,震动研磨===", drop: Ext.encode({data: "喷砂,抛丸,震动研磨"})};
    var ddlCold = {bgc: colorDdl, ta: "center", data: "===冷喷.热烧===", drop: Ext.encode({data: "冷冻去除毛刺,热能去除毛刺"})};
    var ddlStep9 = {bgc: colorDdl, ta: "center", data: "===皮膜处理===", drop: Ext.encode({data: "烤漆前清洗皮膜,一般清洗加皮膜,特殊要求皮膜,清洗,特殊清洗"})};
    var ddlMaching = {bgc: colorDdl, ta: "center", data: "===机加工===", drop: Ext.encode({data: "CNC机加工,传统机加工"})};

//var colorStep="#E6E6E6";
//var colorStepEnd="#A9BCF5";


    function mergeJSON(source1, source2) {
        /*
         * Properties from the Souce1 object will be copied to Source2 Object.
         * Note: This method will return a new merged object, Source1 and Source2 original values will not be replaced.
         * */
        var mergedJSON = Object.create(source2); // Copying Source2 to a new Object

        for (var attrname in source1) {
            if (mergedJSON.hasOwnProperty(attrname)) {
                if (source1[attrname] != null && source1[attrname].constructor == Object) {
                    /*
                     * Recursive call if the property is an object,
                     * Iterate the object and set all properties of the inner object.
                     */
                    mergedJSON[attrname] = mergeJSON(source1[attrname], mergedJSON[attrname]);
                }

            } else {//else copy the property from source1
                mergedJSON[attrname] = source1[attrname];
            }
        }

        return mergedJSON;
    }

    function styleSubTotal(source2) {
        var subtotal = {bgc: colorStepEnd, fm: "money|¥|2|none", dsd: "ed", cal: true};
        return    mergeJSON(subtotal, source2);
    }
    function styleInput(source2) {
        var basic = {bgc: colorInput, ta: "right"};
        return    mergeJSON(basic, source2);
    }

    function setNaToZero(str) {
        return   "=IF(ISNA(" + str + "),0,(" + str + "))";
    }
    function setNaToEmpty(str) {
        return   "=IF(ISNA(" + str + "),,(" + str + "))";
    }



    var json = {
        fileName: "RFQ",
        sheets: [{id: 1, name: "RFQ001a", actived: true, color: "orange"},
            {id: 2, name: "LOOKUP", actived: true, color: "blue"},
            {id: 3, name: "LOOKUP2", actived: true, color: "blue"},
            {id: 5, name: "LOOKUP04", actived: true, color: "blue"},
            {id: 9, name: "Remark", actived: true, color: "green"},
        ],
//        floatings: [
//            {sheet: 1, name: "colGroups", ftype: "colgroup", json: "[{level:1, span:[2,3]}, {level:1, span:[4,6]}]"},
//        ],
        cells: [
            {sheet: 9, row: 0, col: 0, json: {height: 20, va: "middle"}},
            {sheet: 9, row: 0, col: 1, json: {ta: "center", data: "A", width: 50}},
            {sheet: 9, row: 0, col: 2, json: {width: 450}},
//            {sheet: 9, row: 0, col: 3, json: {width: 450}},
            {sheet: 9, row: 1, col: 1, json: {ta: 'center', data: 'Item'}},
            {sheet: 9, row: 1, col: 2, json: {ta: 'center', data: 'Description'}},
            {sheet: 9, row: 2, col: 1, json: {ta: 'center', data: '1'}},
            {sheet: 9, row: 2, col: 2, json: {data: '標示出報價人員需要輸入的部份'}},
            {sheet: 9, row: 3, col: 2, json: {ti: "32px", data: '先實施下拉部份'}},
            {sheet: 9, row: 4, col: 2, json: {ti: "32px", data: '實施填入數值的部份'}},
            {sheet: 9, row: 5, col: 1, json: {ta: 'center', data: '2'}},
            {sheet: 9, row: 5, col: 2, json: {ti: "0px", data: '讓小計默認值為0'}}, //气密性测试 ： 
            {sheet: 9, row: 6, col: 1, json: {ta: 'center', data: '3'}},
            {sheet: 9, row: 6, col: 2, json: {ti: "0px", data: '計算气密性测试，採用LOOKUP04查表取率費'}}, //气密性测试 ： 
            {sheet: 9, row: 7, col: 1, json: {ta: 'center', data: '4'}},
            {sheet: 9, row: 7, col: 2, json: {ti: "0px", data: '計算筛选包装，採用LOOKUP04查表取率費。注意︰&改為中文'}}, //气密性测试 ： 
            {sheet: 9, row: 8, col: 1, json: {ta: 'center', data: ''}},
            {sheet: 9, row: 8, col: 2, json: {ti: "0px", data: '--- temp001c ---'}}, //气密性测试 ： 
            {sheet: 9, row: 9, col: 1, json: {ta: 'center', data: '1'}},
            {sheet: 9, row: 9, col: 2, json: {ti: "0px", data: 'meeting 5月9日(周一)下午，压铸人工费率/H '}}, //气密性测试 ： 
            {sheet: 9, row: 10, col: 2, json: {ti: "0px", data: '--- temp001d 5/11 18:39 ---'}}, //气密性测试 ： 
            {sheet: 9, row: 11, col: 1, json: {ta: 'center', data: '1'}},
            {sheet: 9, row: 11, col: 2, json: {ti: "0px", data: '導出含基本格式及方格底色 '}}, //气密性测试 ： 
           {sheet: 9, row: 12, col: 2, json: {ti: "0px", data: '--- to make real Excel file 5/19 11:44 ---'}}, //气密性测试 ： 
            {sheet: 9, row: 13, col: 1, json: {ta: 'center', data: '1'}},
            {sheet: 9, row: 13, col: 2, json: {ti: "0px", data: '用戶要求要有EXCEL公式'}}, //气密性测试 ： 
            //
            //
            //筛选&包装
            {sheet: 3, row: 0, col: 0, json: {ta: "center", height: 20, va: "middle"}},
            {sheet: 3, row: 0, col: 1, json: {ta: "center", data: "A", width: 200}},
            {sheet: 3, row: 0, col: 2, json: {fm: "money|¥|2|none", ta: "center", data: "B", width: 100}},
            {sheet: 3, row: 0, col: 3, json: {fm: "money|¥|2|none", ta: "center", data: "C", width: 100}},
            {sheet: 3, row: 1, col: 1, json: {ta: 'center', data: '压铸合模费：'}},
            {sheet: 3, row: 2, col: 1, json: {ta: 'center', data: '一。铝合金压铸合模费：'}},
            {sheet: 3, row: 3, col: 1, json: {ta: 'center', data: '铝-125T'}},
            {sheet: 3, row: 4, col: 1, json: {ta: 'center', data: '鋁-150T/160T/200T'}},
            {sheet: 3, row: 5, col: 1, json: {ta: 'center', data: '鋁-250T/280T'}},
            {sheet: 3, row: 6, col: 1, json: {ta: 'center', data: '鋁-350T/340T/400T'}},
            {sheet: 3, row: 7, col: 1, json: {ta: 'center', data: '鋁-550T/530T/500T'}},
            {sheet: 3, row: 8, col: 1, json: {ta: 'center', data: '鋁-630T/650T'}},
            {sheet: 3, row: 9, col: 1, json: {ta: 'center', data: '鋁-800T/900T'}},
            {sheet: 3, row: 10, col: 1, json: {ta: 'center', data: '鋁-1250T'}},
            {sheet: 3, row: 11, col: 1, json: {ta: 'center', data: '鋁-1600T'}},
            {sheet: 3, row: 12, col: 1, json: {ta: 'center', data: '鋁-2000T'}},
            {sheet: 3, row: 13, col: 1, json: {ta: 'center', data: '鋁-2500T'}},
            {sheet: 3, row: 14, col: 1, json: {ta: 'center', data: '鋁-3000T'}},
            {sheet: 3, row: 15, col: 1, json: {ta: 'center', data: '二。锌合金压铸合模费：'}},
            {sheet: 3, row: 16, col: 1, json: {ta: 'center', data: '鋅-快速机/4轴机'}},
            {sheet: 3, row: 17, col: 1, json: {ta: 'center', data: '鋅-10T/5T'}},
            {sheet: 3, row: 18, col: 1, json: {ta: 'center', data: '鋅-15T/20T'}},
            {sheet: 3, row: 19, col: 1, json: {ta: 'center', data: '鋅-25T/30T/40T'}},
            {sheet: 3, row: 20, col: 1, json: {ta: 'center', data: '鋅-50T/60T'}},
            {sheet: 3, row: 21, col: 1, json: {ta: 'center', data: '鋅-80T/100T/125'}},
            {sheet: 3, row: 22, col: 1, json: {ta: 'center', data: '鋅-150T'}},
            {sheet: 3, row: 23, col: 1, json: {ta: 'center', data: '鋅-185/200T'}},
            {sheet: 3, row: 24, col: 1, json: {ta: 'center', data: '鋅-250T'}},
            {sheet: 3, row: 25, col: 1, json: {ta: 'center', data: '鋅-300T'}},
            {sheet: 3, row: 26, col: 1, json: {ta: 'center', data: '三。镁合金压铸合模费：'}},
            {sheet: 3, row: 27, col: 1, json: {ta: 'center', data: '鎂-150T'}},
            {sheet: 3, row: 28, col: 1, json: {ta: 'center', data: '鎂-340T/400T'}},
            {sheet: 3, row: 29, col: 1, json: {ta: 'center', data: '鎂-530T/660T'}},
            {sheet: 3, row: 30, col: 1, json: {ta: 'center', data: 'THX-鎂280T'}},
            {sheet: 3, row: 31, col: 1, json: {ta: 'center', data: 'THX-鎂650T'}},
            {sheet: 3, row: 1, col: 2, json: {ta: 'center', data: '壓鑄機台費率/H'}},
            {sheet: 3, row: 2, col: 2, json: {ta: 'center', data: ''}},
            {sheet: 3, row: 3, col: 2, json: {ta: 'center', data: '150'}},
            {sheet: 3, row: 4, col: 2, json: {ta: 'center', data: '165'}},
            {sheet: 3, row: 5, col: 2, json: {ta: 'center', data: '185'}},
            {sheet: 3, row: 6, col: 2, json: {ta: 'center', data: '210'}},
            {sheet: 3, row: 7, col: 2, json: {ta: 'center', data: '260'}},
            {sheet: 3, row: 8, col: 2, json: {ta: 'center', data: '320'}},
            {sheet: 3, row: 9, col: 2, json: {ta: 'center', data: '430'}},
            {sheet: 3, row: 10, col: 2, json: {ta: 'center', data: '500'}},
            {sheet: 3, row: 11, col: 2, json: {ta: 'center', data: '715'}},
            {sheet: 3, row: 12, col: 2, json: {ta: 'center', data: '915'}},
            {sheet: 3, row: 13, col: 2, json: {ta: 'center', data: '990'}},
            {sheet: 3, row: 14, col: 2, json: {ta: 'center', data: '1110'}},
            {sheet: 3, row: 15, col: 2, json: {ta: 'center', data: ''}},
            {sheet: 3, row: 16, col: 2, json: {ta: 'center', data: '85'}},
            {sheet: 3, row: 17, col: 2, json: {ta: 'center', data: '60'}},
            {sheet: 3, row: 18, col: 2, json: {ta: 'center', data: '80'}},
            {sheet: 3, row: 19, col: 2, json: {ta: 'center', data: '120'}},
            {sheet: 3, row: 20, col: 2, json: {ta: 'center', data: '135'}},
            {sheet: 3, row: 21, col: 2, json: {ta: 'center', data: '160'}},
            {sheet: 3, row: 22, col: 2, json: {ta: 'center', data: '170'}},
            {sheet: 3, row: 23, col: 2, json: {ta: 'center', data: '185'}},
            {sheet: 3, row: 24, col: 2, json: {ta: 'center', data: '200'}},
            {sheet: 3, row: 25, col: 2, json: {ta: 'center', data: '215'}},
            {sheet: 3, row: 26, col: 2, json: {ta: 'center', data: ''}},
            {sheet: 3, row: 27, col: 2, json: {ta: 'center', data: '275'}},
            {sheet: 3, row: 28, col: 2, json: {ta: 'center', data: '350'}},
            {sheet: 3, row: 29, col: 2, json: {ta: 'center', data: '430'}},
            {sheet: 3, row: 30, col: 2, json: {ta: 'center', data: '320'}},
            {sheet: 3, row: 31, col: 2, json: {ta: 'center', data: '430'}},
            {sheet: 3, row: 1, col: 3, json: {ta: 'center', data: '人工費率/H'}},
            {sheet: 3, row: 2, col: 3, json: {ta: 'center', data: ''}},
            {sheet: 3, row: 3, col: 3, json: {ta: 'center', data: '40.9'}},
            {sheet: 3, row: 4, col: 3, json: {ta: 'center', data: '40.9'}},
            {sheet: 3, row: 5, col: 3, json: {ta: 'center', data: '40.9'}},
            {sheet: 3, row: 6, col: 3, json: {ta: 'center', data: '40.9'}},
            {sheet: 3, row: 7, col: 3, json: {ta: 'center', data: '56.3'}},
            {sheet: 3, row: 8, col: 3, json: {ta: 'center', data: '56.3'}},
            {sheet: 3, row: 9, col: 3, json: {ta: 'center', data: '71.6'}},
            {sheet: 3, row: 10, col: 3, json: {ta: 'center', data: '92'}},
            {sheet: 3, row: 11, col: 3, json: {ta: 'center', data: '102.3'}},
            {sheet: 3, row: 12, col: 3, json: {ta: 'center', data: '102.3'}},
            {sheet: 3, row: 13, col: 3, json: {ta: 'center', data: '102.3'}},
            {sheet: 3, row: 14, col: 3, json: {ta: 'center', data: '102.3'}},
            {sheet: 3, row: 15, col: 3, json: {ta: 'center', data: ''}},
            {sheet: 3, row: 16, col: 3, json: {ta: 'center', data: '11.3'}},
            {sheet: 3, row: 17, col: 3, json: {ta: 'center', data: '11.3'}},
            {sheet: 3, row: 18, col: 3, json: {ta: 'center', data: '16.9'}},
            {sheet: 3, row: 19, col: 3, json: {ta: 'center', data: '16.9'}},
            {sheet: 3, row: 20, col: 3, json: {ta: 'center', data: '25.3'}},
            {sheet: 3, row: 21, col: 3, json: {ta: 'center', data: '33.8'}},
            {sheet: 3, row: 22, col: 3, json: {ta: 'center', data: '33.8'}},
            {sheet: 3, row: 23, col: 3, json: {ta: 'center', data: '33.8'}},
            {sheet: 3, row: 24, col: 3, json: {ta: 'center', data: '33.8'}},
            {sheet: 3, row: 25, col: 3, json: {ta: 'center', data: '33.8'}},
            {sheet: 3, row: 26, col: 3, json: {ta: 'center', data: ''}},
            {sheet: 3, row: 27, col: 3, json: {ta: 'center', data: '101.3'}},
            {sheet: 3, row: 28, col: 3, json: {ta: 'center', data: '112.5'}},
            {sheet: 3, row: 29, col: 3, json: {ta: 'center', data: '112.5'}},
            {sheet: 3, row: 30, col: 3, json: {ta: 'center', data: '112.5'}},
            {sheet: 3, row: 31, col: 3, json: {ta: 'center', data: '112.5'}},
            {sheet: 5, row: 0, col: 0, json: {dsd: "", height: 20, va: "middle"}},
            {sheet: 5, row: 0, col: 1, json: {ta: "center", data: "A", width: 200}},
            {sheet: 5, row: 0, col: 2, json: {fm: "money|¥|2|none", ta: "center", data: "B", width: 100}},
            {sheet: 5, row: 1, col: 1, json: {data: '後段加工 工时费率：'}},
            {sheet: 5, row: 2, col: 1, json: {data: ''}},
            {sheet: 5, row: 3, col: 1, json: {data: 'CNC机加工'}},
            {sheet: 5, row: 4, col: 1, json: {data: '传统机加工'}},
            {sheet: 5, row: 5, col: 1, json: {data: '毛刺处理费（含粗磨） '}},
            {sheet: 5, row: 6, col: 1, json: {data: '打磨'}},
            {sheet: 5, row: 7, col: 1, json: {data: ' 整形'}},
            {sheet: 5, row: 8, col: 1, json: {data: '毛刺處理'}},
            {sheet: 5, row: 9, col: 1, json: {data: '镜面抛光'}},
            {sheet: 5, row: 10, col: 1, json: {data: '震动研磨'}},
            {sheet: 5, row: 11, col: 1, json: {data: '抛丸'}},
            {sheet: 5, row: 12, col: 1, json: {data: '喷砂'}},
            {sheet: 5, row: 13, col: 1, json: {data: '热能去除毛刺'}},
            {sheet: 5, row: 14, col: 1, json: {data: '冷冻去除毛刺'}},
            {sheet: 5, row: 15, col: 1, json: {data: '气密性测试'}},
            {sheet: 5, row: 16, col: 1, json: {data: '丝印/网印'}},
            {sheet: 5, row: 17, col: 1, json: {data: '筛选和包装'}},
            {sheet: 5, row: 1, col: 1, json: {data: '費率/H '}},
            {sheet: 5, row: 1, col: 2, json: {data: '費率/H '}},
            {sheet: 5, row: 2, col: 2, json: {data: ''}},
            {sheet: 5, row: 3, col: 2, json: {data: '66'}},
            {sheet: 5, row: 4, col: 2, json: {data: '45'}},
            {sheet: 5, row: 5, col: 2, json: {data: '35'}},
            {sheet: 5, row: 6, col: 2, json: {data: '45'}},
            {sheet: 5, row: 7, col: 2, json: {data: '45'}},
            {sheet: 5, row: 8, col: 2, json: {data: '35'}},
            {sheet: 5, row: 9, col: 2, json: {data: '60'}},
            {sheet: 5, row: 10, col: 2, json: {data: '120'}},
            {sheet: 5, row: 11, col: 2, json: {data: '85'}},
            {sheet: 5, row: 12, col: 2, json: {data: '125'}},
            {sheet: 5, row: 13, col: 2, json: {data: '206'}},
            {sheet: 5, row: 14, col: 2, json: {data: '115'}},
            {sheet: 5, row: 15, col: 2, json: {data: '45'}},
            {sheet: 5, row: 16, col: 2, json: {data: '222'}},
            {sheet: 5, row: 17, col: 2, json: {data: '35'}},
            {sheet: 2, row: 0, col: 0, json: {ta: "center", height: 20, va: "middle"}},
            {sheet: 2, row: 0, col: 1, json: {data: "A", width: 200}},
            {sheet: 2, row: 0, col: 2, json: {fm: "money|¥|2|none", data: "B", width: 100}},
            {sheet: 2, row: 0, col: 3, json: {fm: "money|¥|2|none", data: "C", width: 100}},
            {sheet: 2, row: 0, col: 4, json: {data: "D", width: 200}},
            {sheet: 2, row: 0, col: 5, json: {data: "E", width: 100}},
            {sheet: 2, row: 1, col: 1, json: {dsd: "ed", data: "材料規格"}},
            {sheet: 2, row: 1, col: 2, json: {dsd: "ed", data: "材料價格/kg"}},
            {sheet: 2, row: 1, col: 3, json: {dsd: "ed", data: "廢料價格/kg"}},
            {sheet: 2, row: 2, col: 1, json: {dsd: "ed", data: "一。铝合金"}},
            {sheet: 2, row: 2, col: 2, json: {dsd: "ed", data: ""}},
            {sheet: 2, row: 2, col: 3, json: {dsd: "ed", data: ""}},
            {sheet: 2, row: 3, col: 1, json: {dsd: "ed", data: "AlSi10Mg"}},
            {sheet: 2, row: 3, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 3, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 4, col: 1, json: {dsd: "ed", data: "AlSi9Cu3"}},
            {sheet: 2, row: 4, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 4, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 5, col: 1, json: {dsd: "ed", data: "ADC12"}},
            {sheet: 2, row: 5, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 5, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 6, col: 1, json: {dsd: "ed", data: "A380"}},
            {sheet: 2, row: 6, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 6, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 7, col: 1, json: {dsd: "ed", data: "A413"}},
            {sheet: 2, row: 7, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 7, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 8, col: 1, json: {dsd: "ed", data: "A360"}},
            {sheet: 2, row: 8, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 8, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 9, col: 1, json: {dsd: "ed", data: "A356"}},
            {sheet: 2, row: 9, col: 2, json: {dsd: "ed", data: "13.25"}},
            {sheet: 2, row: 9, col: 3, json: {dsd: "ed", data: "11.25"}},
            {sheet: 2, row: 10, col: 1, json: {dsd: "ed", data: " 二。锌合金"}},
            {sheet: 2, row: 10, col: 2, json: {dsd: "ed", data: ""}},
            {sheet: 2, row: 10, col: 3, json: {dsd: "ed", data: ""}},
            {sheet: 2, row: 11, col: 1, json: {dsd: "ed", data: "ZINC-2"}},
            {sheet: 2, row: 11, col: 2, json: {dsd: "ed", data: "15.81"}},
            {sheet: 2, row: 11, col: 3, json: {dsd: "ed", data: "12.81"}},
            {sheet: 2, row: 12, col: 1, json: {dsd: "ed", data: "ZINC-3"}},
            {sheet: 2, row: 12, col: 2, json: {dsd: "ed", data: "15.81"}},
            {sheet: 2, row: 12, col: 3, json: {dsd: "ed", data: "12.81"}},
            {sheet: 2, row: 13, col: 1, json: {dsd: "ed", data: "ZINC-5"}},
            {sheet: 2, row: 13, col: 2, json: {dsd: "ed", data: "15.81"}},
            {sheet: 2, row: 13, col: 3, json: {dsd: "ed", data: "12.81"}},
            {sheet: 2, row: 14, col: 1, json: {dsd: "ed", data: "ZINC-7"}},
            {sheet: 2, row: 14, col: 2, json: {dsd: "ed", data: "16.67"}},
            {sheet: 2, row: 14, col: 3, json: {dsd: "ed", data: "13.67"}},
            {sheet: 2, row: 15, col: 1, json: {dsd: "ed", data: "ZAMAK-8"}},
            {sheet: 2, row: 15, col: 2, json: {dsd: "ed", data: "16.67"}},
            {sheet: 2, row: 15, col: 3, json: {dsd: "ed", data: "13.67"}},
            {sheet: 2, row: 16, col: 1, json: {dsd: "ed", data: " 三。镁合金"}},
            {sheet: 2, row: 16, col: 2, json: {dsd: "ed", data: ""}},
            {sheet: 2, row: 16, col: 3, json: {dsd: "ed", data: ""}},
            {sheet: 2, row: 17, col: 1, json: {dsd: "ed", data: "AZ91D"}},
            {sheet: 2, row: 17, col: 2, json: {dsd: "ed", data: "17.09"}},
            {sheet: 2, row: 17, col: 3, json: {dsd: "ed", data: "14.09"}},
            {sheet: 2, row: 18, col: 1, json: {dsd: "ed", data: "AM60"}},
            {sheet: 2, row: 18, col: 2, json: {dsd: "ed", data: "16.24"}},
            {sheet: 2, row: 18, col: 3, json: {dsd: "ed", data: "13.24"}},
            {sheet: 2, row: 19, col: 1, json: {dsd: "ed", data: "THX-AZ91D"}},
            {sheet: 2, row: 19, col: 2, json: {dsd: "ed", data: "22.29"}},
            {sheet: 2, row: 19, col: 3, json: {dsd: "ed", data: "14.09"}},
            {sheet: 2, row: 20, col: 1, json: {dsd: "ed", data: "THX-AM60"}},
            {sheet: 2, row: 20, col: 2, json: {dsd: "ed", data: "22.29"}},
            {sheet: 2, row: 20, col: 3, json: {dsd: "ed", data: "14.09"}},
            {sheet: 1, row: 0, col: 0, json: {dsd: "", height: 20, va: "middle"}},
            {sheet: 1, row: 0, col: 1, json: {ta: "center", dsd: "", data: "A", width: 50}},
            {sheet: 1, row: 0, col: 2, json: {ta: "center", data: "B", width: 200}},
            {sheet: 1, row: 0, col: 3, json: {dsd: "", width: 200}},
            {sheet: 1, row: 0, col: 4, json: {dsd: "", width: 200}},
            {sheet: 1, row: 0, col: 5, json: {dsd: "", width: 200}},
            {sheet: 1, row: 0, col: 6, json: {dsd: "", width: 200}},
            {sheet: 1, row: 0, col: 7, json: {dsd: "", width: 200}},
            {sheet: 1, row: 0, col: 8, json: {dsd: "", width: 200}},
            {sheet: 1, row: 1, col: 0, json: {hidden: true}},
            {sheet: 1, row: 2, col: 0, json: {hidden: true}},
//            {sheet: 1, row: 1, col: 1, json: {data: ""}},
//            {sheet: 1, row: 1, col: 2, json: {data: ""}},
//            {sheet: 1, row: 1, col: 3, json: {data: ""}},
//            {sheet: 1, row: 2, col: 1, json: {data: ""}},
//            {sheet: 1, row: 2, col: 2, json: {data: ""}},
//                    {sheet: 1, row: 2, col: 3, json: {data: "0"}},
//            {sheet: 1, row: 2, col: 3, json: {data: "TODO 保存數據", it: "button", btnStyle: "color: #900; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK_CALLBACK_FN"}},
//   {sheet: 1, row: 2, col: 3, json: {data: "TODO 保存數據", it: "button", btnStyle: "color: #900; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK_CALLBACK_FN"}},

//           {sheet: 1, row: 3, col: 1, json: {data: "EXP1", it: "button", btnStyle: "color: #900; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK_CALLBACK_FN"}},
            {sheet: 1, row: 3, col: 2, json: {data: "接受询价日期："}},
            //  {sheet: 1, row: 2, col: 3, json:{ data: "TODO 保存數據", it: "button", btnStyle: "color: #900; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK_CALLBACK_FN" } },
            {sheet: 1, row: 4, col: 1, json: {data: "生成", it: "button", btnStyle: "color: blue; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK___MAKE_EXCEL"}},
            {sheet: 1, row: 4, col: 2, json: {data: "业务担当："}},
            {sheet: 1, row: 4, col: 3, json: {data: ""}},
            {sheet: 1, row: 5, col: 1, json: {data: ""}},
            {sheet: 1, row: 5, col: 2, json: {data: ""}},
            {sheet: 1, row: 5, col: 3, json: {data: ""}},
            {sheet: 1, row: 6, col: 1, json: {data: ""}},
            {sheet: 1, row: 6, col: 2, json: {data: "产品简图 "}},
            {sheet: 1, row: 6, col: 3, json: {data: ""}},
            {sheet: 1, row: 7, col: 1, json: {data: ""}},
            {sheet: 1, row: 7, col: 2, json: {data: "项目 料号 "}},
            {sheet: 1, row: 7, col: 3, json: {ta: "center", data: "(1)12-00620-00003_R6"}},
            {sheet: 1, row: 8, col: 1, json: {data: ""}},
            {sheet: 1, row: 8, col: 2, json: {data: "品名 / 图纸版本号："}},
            {sheet: 1, row: 8, col: 3, json: {ta: "center", data: "REV 6"}},
            {sheet: 1, row: 9, col: 1, json: {data: ""}},
            {sheet: 1, row: 9, col: 2, json: {data: "产品外形尺寸"}},
            {sheet: 1, row: 9, col: 3, json: {ta: "center", data: "41.25*29*29"}},
            {sheet: 1, row: 10, col: 1, json: {data: ""}},
            {sheet: 1, row: 10, col: 2, json: {data: "材质规格： "}},
            {sheet: 1, row: 10, col: 3, json: ddlMaterial},
            {sheet: 1, row: 11, col: 1, json: {data: ""}},
            {sheet: 1, row: 11, col: 2, json: {data: "产品单重（g）："}},
            {sheet: 1, row: 11, col: 3, json: styleInput({fm: "number", data: "13"})},
            {sheet: 1, row: 12, col: 1, json: {data: ""}},
            {sheet: 1, row: 12, col: 2, json: {data: "表面要求： "}},
            {sheet: 1, row: 12, col: 3, json: ddlSurface},
            {sheet: 1, row: 13, col: 1, json: {data: ""}},
            {sheet: 1, row: 13, col: 2, json: {data: "年需求量："}},
            {sheet: 1, row: 13, col: 3, json: {ta: "center", data: "12000"}},
            {sheet: 1, row: 14, col: 1, json: {data: ""}},
            {sheet: 1, row: 14, col: 2, json: {data: ""}},
            {sheet: 1, row: 14, col: 3, json: {data: ""}},
            {sheet: 1, row: 15, col: 1, json: {data: ""}},
            {sheet: 1, row: 15, col: 2, json: {bgc: colorStep, data: "模具费用： "}},
            {sheet: 1, row: 15, col: 3, json: {data: ""}},
            {sheet: 1, row: 16, col: 1, json: {data: ""}},
            {sheet: 1, row: 16, col: 2, json: {data: "模穴数 "}},
            {sheet: 1, row: 16, col: 3, json: styleInput({fm: "number", data: "1"})},
            {sheet: 1, row: 17, col: 1, json: {data: ""}},
            {sheet: 1, row: 17, col: 2, json: {data: "侧滑模数/油压抽芯数 "}},
            {sheet: 1, row: 17, col: 3, json: {ta: "center", data: "2"}},
            {sheet: 1, row: 18, col: 1, json: {data: ""}},
            {sheet: 1, row: 18, col: 2, json: {data: "模具寿命："}},
            {sheet: 1, row: 18, col: 3, json: {ta: "center", data: "100K"}},
            {sheet: 1, row: 19, col: 1, json: {data: ""}},
            {sheet: 1, row: 19, col: 2, json: {data: "压铸模费用： "}},
            {sheet: 1, row: 19, col: 3, json: styleInput({fm: "money|¥|2|none", data: "71500"})},
            {sheet: 1, row: 20, col: 1, json: {data: ""}},
            {sheet: 1, row: 20, col: 2, json: {data: "切边模费用： "}},
            {sheet: 1, row: 20, col: 3, json: styleInput({fm: "money|¥|2|none", data: "12000"})},
            {sheet: 1, row: 21, col: 1, json: {data: ""}},
            {sheet: 1, row: 21, col: 2, json: {data: "加工夹治具费用： "}},
            {sheet: 1, row: 21, col: 3, json: styleInput({fm: "money|¥|2|none", data: "8000"})},
            {sheet: 1, row: 22, col: 1, json: {data: ""}},
            {sheet: 1, row: 22, col: 2, json: {data: "烤漆夹治具费用： "}},
            {sheet: 1, row: 22, col: 3, json: styleInput({fm: "money|¥|2|none", data: "8000"})},
            {sheet: 1, row: 23, col: 1, json: {data: ""}},
            {sheet: 1, row: 23, col: 2, json: {bgc: colorStepEnd, data: "模具总价（人民币）： "}},
            {sheet: 1, row: 23, col: 3, json: styleSubTotal({data: "=(C19+C20+C21+C22)"})},
            {sheet: 1, row: 24, col: 1, json: {data: ""}},
            {sheet: 1, row: 24, col: 2, json: {data: "模具总价（USD）： "}},
            {sheet: 1, row: 24, col: 3, json: {fm: "money|$|2|none", dsd: "ed", cal: true, data: "=(C23/6.35)"}},
            {sheet: 1, row: 25, col: 0, json: {hidden: true}},
            {sheet: 1, row: 26, col: 0, json: {hidden: true}},
//          
//            {sheet: 1, row: 25, col: 1, json: {data: ""}},
//            {sheet: 1, row: 25, col: 2, json: {data: ""}},
//            {sheet: 1, row: 25, col: 3, json: {data: ""}},
//            {sheet: 1, row: 26, col: 1, json: {data: ""}},
//            {sheet: 1, row: 26, col: 2, json: {data: ""}},
//            {sheet: 1, row: 26, col: 3, json: {data: ""}},
//            
            {sheet: 1, row: 27, col: 1, json: {data: ""}},
            {sheet: 1, row: 27, col: 2, json: {data: "产品单价： "}},
            {sheet: 1, row: 27, col: 3, json: {data: ""}},
            {sheet: 1, row: 28, col: 1, json: {data: "* "}},
            {sheet: 1, row: 28, col: 2, json: {bgc: colorStep, data: "1)材料费 ： "}},
            {sheet: 1, row: 28, col: 3, json: {data: ""}},
            {sheet: 1, row: 29, col: 1, json: {data: "1-1）"}},
            {sheet: 1, row: 29, col: 2, json: {data: "材质规格："}},
            {sheet: 1, row: 29, col: 3, json: {ta: "center", dsd: "ed", data: "=C10", cal: true}},
            {sheet: 1, row: 30, col: 1, json: {data: "1-2）"}},
            {sheet: 1, row: 30, col: 2, json: {data: "材料单价/KG :"}},
            {sheet: 1, row: 30, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C10,LOOKUP!$A$3:$C$20,2,0)"}},
            {sheet: 1, row: 31, col: 1, json: {data: "1-3）"}},
            {sheet: 1, row: 31, col: 2, json: {data: "单重（g）："}},
            {sheet: 1, row: 31, col: 3, json: {dsd: "ed", data: "=C11", cal: true}},
            {sheet: 1, row: 32, col: 1, json: {data: "1-4）"}},
            {sheet: 1, row: 32, col: 2, json: {data: "产品材料费（净重价格）： "}},
            {sheet: 1, row: 32, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", data: "=C30*C31/1000", cal: true}},
            {sheet: 1, row: 33, col: 1, json: {data: "1-5）"}},
            {sheet: 1, row: 33, col: 2, json: {data: "料柄流道渣包等废料重量(g)： "}},
            {sheet: 1, row: 33, col: 3, json: styleInput({fm: "number", data: "350"})},
            {sheet: 1, row: 34, col: 1, json: {data: "1-6）"}},
            {sheet: 1, row: 34, col: 2, json: {data: "料柄流道渣包比率 ： "}},
            {sheet: 1, row: 34, col: 3, json: {fm: "number", dfm: "0%", dsd: "ed", cal: true, data: "=100*C31*C16/(C31*C16+C33)"}},
            {sheet: 1, row: 35, col: 1, json: {data: "1-7）"}},
            {sheet: 1, row: 35, col: 2, json: {data: "料柄流道渣包等废料价格/Kg： "}},
            {sheet: 1, row: 35, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C10,LOOKUP!$A$3:$C$20,3,0)"}},
            {sheet: 1, row: 36, col: 1, json: {data: "1-8）"}},
            {sheet: 1, row: 36, col: 2, json: {data: "料柄流道渣包回收价差损失额 ： "}},
            {sheet: 1, row: 36, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=(C30-C35)*C33/1000/C16"}},
            {sheet: 1, row: 37, col: 1, json: {data: "1-9）"}},
            {sheet: 1, row: 37, col: 2, json: {data: "压铸熔炼材料氧化损耗(量）2% "}},
            {sheet: 1, row: 37, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=(C31+C33)*C30*0.02/1000/C16"}},
            {sheet: 1, row: 38, col: 1, json: {data: ""}},
            {sheet: 1, row: 38, col: 2, json: {bgc: colorStepEnd, data: "材料费　小计： "}},
            {sheet: 1, row: 38, col: 3, json: styleSubTotal({data: "=IF(ISNA(C32+C36+C37),0,(C32+C36+C37))"})},
            {sheet: 1, row: 39, col: 1, json: {data: "* "}},
            {sheet: 1, row: 39, col: 2, json: {bgc: colorStep, data: "2)压铸费(含切边）： "}},
            {sheet: 1, row: 39, col: 3, json: {ta: 'center', data: ""}},
            {sheet: 1, row: 40, col: 1, json: {data: "2-1）"}},
            {sheet: 1, row: 40, col: 2, json: {data: "适用机型： "}},
            {sheet: 1, row: 40, col: 3, json: ddlMachine},
            {sheet: 1, row: 41, col: 1, json: {data: "2-2）"}},
            {sheet: 1, row: 41, col: 2, json: {data: "设备费率/H： "}},
            {sheet: 1, row: 41, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C40,LOOKUP2!$A$1:$C$100,2,0)"}},
            {sheet: 1, row: 42, col: 1, json: {data: "2-3）"}},
            {sheet: 1, row: 42, col: 2, json: {data: "工时（秒）/只 ："}},
            {sheet: 1, row: 42, col: 3, json: styleInput({fm: "number", data: "30"})},
            {sheet: 1, row: 43, col: 1, json: {data: "2-4）"}},
            {sheet: 1, row: 43, col: 2, json: {data: "产能 只/H ："}},
            {sheet: 1, row: 43, col: 3, json: {fm: "number", dfm: "0", dsd: "ed", cal: true, data: "=3600/C42"}},
            {sheet: 1, row: 44, col: 1, json: {data: "2-5）"}},
            {sheet: 1, row: 44, col: 2, json: {data: "设备费/只： "}},
            {sheet: 1, row: 44, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=C41/C43"}},
            {sheet: 1, row: 45, col: 1, json: {data: "2-6）"}},
            {sheet: 1, row: 45, col: 2, json: {data: "人工费/只 ： "}},
            {sheet: 1, row: 45, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C40,LOOKUP2!$A$1:$C$100,3,0)/C43"}},
//            {sheet: 1, row: 46, col: 1, json: {data: ""}},
//            {sheet: 1, row: 46, col: 2, json: {data: ""}},
//            {sheet: 1, row: 46, col: 3, json: {data: ""}},
            {sheet: 1, row: 46, col: 1, json: {data: "2-7）"}},
            {sheet: 1, row: 46, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 46, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C40,LOOKUP2!$A$1:$C$100,3,0)"}},
            {sheet: 1, row: 47, col: 1, json: {data: "2-8）"}},
            {sheet: 1, row: 47, col: 2, json: {data: "良品率： "}},
            {sheet: 1, row: 47, col: 3, json: styleInput({fm: "number", dfm: "0%", data: "95"})},
            {sheet: 1, row: 48, col: 1, json: {data: ""}},
            {sheet: 1, row: 48, col: 2, json: {bgc: colorStepEnd, data: "压铸费　小计： "}},
            {sheet: 1, row: 48, col: 3, json: styleSubTotal({data: "=(C44+C45)*(1+(1-C47/100))/C16"})},
            {sheet: 1, row: 48, col: 3, json: styleSubTotal({data: setNaToZero("(C44+C45)*(1+(1-C47/100))/C16")})},
//            {sheet: 1, row: 48, col: 3, json: styleSubTotal({data: setNaToZero("(C44 + C45) * (1 + (1 - C47 / 100)) / C16)")})},
            {sheet: 1, row: 49, col: 1, json: {data: "* "}},
            {sheet: 1, row: 49, col: 2, json: {bgc: colorStep, data: "3)毛刺处理费（含粗磨）"}},
            {sheet: 1, row: 49, col: 3, json: {data: ""}},
            {sheet: 1, row: 50, col: 1, json: {data: "3-1）"}},
            {sheet: 1, row: 50, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 50, col: 3, json: styleInput({fm: "number", data: "82"})},
            {sheet: 1, row: 51, col: 1, json: {data: "3-2）"}},
            {sheet: 1, row: 51, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 51, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", data: "35"}},
            {sheet: 1, row: 52, col: 1, json: {data: ""}},
            {sheet: 1, row: 52, col: 2, json: {bgc: colorStepEnd, data: "毛刺处理费　小计： "}},
            {sheet: 1, row: 52, col: 3, json: styleSubTotal({data: "=(C50/3600)*C51"})},
            {sheet: 1, row: 53, col: 1, json: {data: "* "}},
            {sheet: 1, row: 53, col: 2, json: {bgc: colorStep, data: "4)外观打磨费（入/溢料口, 分模线及一般外观瑕疵等部位做打磨）"}},
            {sheet: 1, row: 53, col: 3, json: {data: ""}},
            {sheet: 1, row: 54, col: 1, json: {data: "4-1）"}},
            {sheet: 1, row: 54, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 54, col: 3, json: styleInput({fm: "number", data: "150"})},
            {sheet: 1, row: 55, col: 1, json: {data: "4-2）"}},
            {sheet: 1, row: 55, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 55, col: 3, json: {dsd: true, data: "45"}},
            {sheet: 1, row: 56, col: 1, json: {data: "4-3）"}},
            {sheet: 1, row: 56, col: 2, json: {data: "打磨费： "}},
            {sheet: 1, row: 56, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=(C54/3600)*C55"}},
            {sheet: 1, row: 57, col: 1, json: {data: "4-4）"}},
            {sheet: 1, row: 57, col: 2, json: {data: "良品率： "}},
            {sheet: 1, row: 57, col: 3, json: styleInput({fm: "number", dfm: "0%", data: "99"})},
            {sheet: 1, row: 58, col: 0, json: {hidden: true}},
//            {sheet: 1, row: 58, col: 2, json: {data: ""}},
//            {sheet: 1, row: 58, col: 3, json: {data: ""}},
            {sheet: 1, row: 59, col: 1, json: {data: ""}},
            {sheet: 1, row: 59, col: 2, json: {bgc: colorStepEnd, data: "外观打磨费　小计："}},
            {sheet: 1, row: 59, col: 3, json: styleSubTotal({data: "=C56*(1+(1-C57/100))"})},
            {sheet: 1, row: 60, col: 1, json: {data: "* "}},
            {sheet: 1, row: 60, col: 2, json: {bgc: colorStep, data: "5)平整度整形费 "}},
            {sheet: 1, row: 60, col: 3, json: {data: ""}},
            {sheet: 1, row: 61, col: 1, json: {data: "5-1）"}},
            {sheet: 1, row: 61, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 61, col: 3, json: styleInput({fm: "number", data: "60"})},
            {sheet: 1, row: 62, col: 1, json: {data: "5-2）"}},
            {sheet: 1, row: 62, col: 2, json: {data: "人工费率/H： "}},
            {sheet: 1, row: 62, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", data: "45"}},
            {sheet: 1, row: 63, col: 1, json: {data: "5-3）"}},
            {sheet: 1, row: 63, col: 2, json: {data: "人工费/只 "}},
            {sheet: 1, row: 63, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=(C61/3600)*C62"}},
            {sheet: 1, row: 64, col: 1, json: {data: ""}},
            {sheet: 1, row: 64, col: 2, json: {bgc: colorStepEnd, data: "平整度整形费　小计： "}},
            {sheet: 1, row: 64, col: 3, json: styleSubTotal({data: "=(C61/3600)*C62"})},
            {sheet: 1, row: 65, col: 1, json: {data: "* "}},
            {sheet: 1, row: 65, col: 2, json: {bgc: colorStep, data: "6)机加工 "}},
            {sheet: 1, row: 65, col: 3, json: ddlMaching},
            {sheet: 1, row: 66, col: 1, json: {data: "6-1）"}},
            {sheet: 1, row: 66, col: 2, json: {data: "机加工工时（秒）/只 ： "}},
            {sheet: 1, row: 66, col: 3, json: {fm: "number", dfm: "0", data: "30"}},
            {sheet: 1, row: 67, col: 1, json: {data: "6-2）"}},
            {sheet: 1, row: 67, col: 2, json: {data: "机加工费率/H： "}},
            {sheet: 1, row: 67, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C65,LOOKUP04!$A$2:$B$17,2,0)"}},
            {sheet: 1, row: 68, col: 1, json: {data: "6-3）"}},
            {sheet: 1, row: 68, col: 2, json: {data: "机加工良品率： "}},
            {sheet: 1, row: 68, col: 3, json: styleInput({fm: "number", dfm: "0%", data: "95"})},
            {sheet: 1, row: 69, col: 1, json: {data: ""}},
            {sheet: 1, row: 69, col: 2, json: {bgc: colorStepEnd, data: "机加工　小计： "}},
//            {sheet: 1, row: 69, col: 3, json: styleSubTotal({data: "=(C66/3600)*C67 * (1 + (1 - C68 / 100))"})},
//            {sheet: 1, row: 69, col: 3, json: styleSubTotal({data: "=IF(ISNA((C66/3600)*C67 * (1 + (1 - C68 / 100))),0,(C66/3600)*C67 * (1 + (1 - C68 / 100)))"})},
            {sheet: 1, row: 69, col: 3, json: styleSubTotal({data: setNaToZero("(C66/3600)*C67 * (1 + (1 - C68 / 100))")})},
            {sheet: 1, row: 70, col: 1, json: {data: "* "}},
            {sheet: 1, row: 70, col: 2, json: {bgc: colorStep, data: "7)冷喷.热烧毛刺 &费率/分： "}},
            {sheet: 1, row: 70, col: 3, json: ddlCold},
            {sheet: 1, row: 71, col: 1, json: {data: "7-1）"}},
            {sheet: 1, row: 71, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 71, col: 3, json: styleInput({fm: "number", dfm: "0", data: "60"})},
            {sheet: 1, row: 72, col: 1, json: {data: "7-2）"}},
            {sheet: 1, row: 72, col: 2, json: {data: "加工费率/H： "}},
            {sheet: 1, row: 72, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C70,LOOKUP04!$A$2:$B$17,2,0)"}},
            {sheet: 1, row: 73, col: 1, json: {data: ""}},
            {sheet: 1, row: 73, col: 2, json: {bgc: colorStepEnd, data: "冷喷热烧毛刺　小计："}},
            {sheet: 1, row: 73, col: 3, json: styleSubTotal({data: setNaToZero("(C71/3600)*C72")})},
            {sheet: 1, row: 74, col: 1, json: {data: "* "}},
            {sheet: 1, row: 74, col: 2, json: {bgc: colorStep, data: "8)喷沙.抛丸.震动研磨 &费率 "}},
            {sheet: 1, row: 74, col: 3, json: ddlSand},
            {sheet: 1, row: 75, col: 1, json: {data: "8-1）"}},
            {sheet: 1, row: 75, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 75, col: 3, json: styleInput({fm: "number", dfm: "0", data: "60"})},
            {sheet: 1, row: 76, col: 1, json: {data: "8-2）"}},
            {sheet: 1, row: 76, col: 2, json: {data: "加工费率/H： "}},
            {sheet: 1, row: 76, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=VLOOKUP(C74,LOOKUP04!$A$2:$B$17,2,0)"}},
            {sheet: 1, row: 77, col: 1, json: {data: ""}},
            {sheet: 1, row: 77, col: 2, json: {bgc: colorStepEnd, data: "喷沙抛丸震动研磨　小计： "}},
//            {sheet: 1, row: 77, col: 3, json: {bgc: colorStepEnd, fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=(C75/3600)*C76"}},
            {sheet: 1, row: 77, col: 3, json: styleSubTotal({data: "=IF(ISNA(C76),0,(C75/3600)*C76)"})},
            {sheet: 1, row: 78, col: 1, json: {data: "* "}},
            {sheet: 1, row: 78, col: 2, json: {bgc: colorStep, data: "9)皮膜处理 &费率 "}},
            {sheet: 1, row: 78, col: 3, json: {data: ""}},
            {sheet: 1, row: 79, col: 1, json: {data: "9-1）"}},
            {sheet: 1, row: 79, col: 2, json: {data: "皮膜处理： "}},
            {sheet: 1, row: 79, col: 3, json: ddlStep9},
            {sheet: 1, row: 80, col: 1, json: {data: "9-2）"}},
            {sheet: 1, row: 80, col: 2, json: {data: "工件表面面积 (算单面）DM2： "}},
            {sheet: 1, row: 80, col: 3, json: styleInput({fm: "number", dfm: "0.00", data: "0.99"})},
            {sheet: 1, row: 81, col: 1, json: {data: "9-3）"}},
            {sheet: 1, row: 81, col: 2, json: {data: "加工费率/H： "}},
            {sheet: 1, row: 81, col: 3, json: styleInput({fm: "money|¥|2|none", cal: true, data: "0.88"})},
            {sheet: 1, row: 82, col: 1, json: {data: ""}},
            {sheet: 1, row: 82, col: 2, json: {data: "難度係數 "}},
            {sheet: 1, row: 82, col: 3, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
            {sheet: 1, row: 83, col: 1, json: {data: ""}},
            {sheet: 1, row: 83, col: 2, json: {bgc: colorStepEnd, data: "皮膜处理　小计： "}},
            {sheet: 1, row: 83, col: 3, json: styleSubTotal({data: "=C80*C81*C82"})},
            {sheet: 1, row: 84, col: 1, json: {data: "* "}},
            {sheet: 1, row: 84, col: 2, json: {bgc: colorStep, data: "10)粉体烤漆.液体烤漆.丝网印： "}},
            {sheet: 1, row: 84, col: 3, json: {data: ""}},
            {sheet: 1, row: 85, col: 1, json: {data: "10-1）"}},
            {sheet: 1, row: 85, col: 2, json: {data: "表面要求"}},
            {sheet: 1, row: 85, col: 3, json: {dsd: "ed", ta: "center", cal: true, data: "=C12"}},
            {sheet: 1, row: 86, col: 1, json: {data: "10-2）"}},
            {sheet: 1, row: 86, col: 2, json: {data: "工件表面积(算需喷漆面积）DM2： "}},
            {sheet: 1, row: 86, col: 3, json: styleInput({fm: "number", data: "6"})},
            {sheet: 1, row: 87, col: 1, json: {data: "10-3）"}},
            {sheet: 1, row: 87, col: 2, json: {data: "漆材料费/dm： "}},
            {sheet: 1, row: 87, col: 3, json: styleInput({fm: "money|¥|2|none", data: "0.1"})},
            {sheet: 1, row: 88, col: 1, json: {data: "10-4）"}},
            {sheet: 1, row: 88, col: 2, json: {data: "烤漆设备+人工費率/dm： "}},
            {sheet: 1, row: 88, col: 3, json: styleInput({fm: "money|¥|2|none", data: "0.6"})},
            {sheet: 1, row: 89, col: 1, json: {data: "10-5）"}},
            {sheet: 1, row: 89, col: 2, json: {data: "良品率： "}},
            {sheet: 1, row: 89, col: 3, json: styleInput({fm: "number", dfm: "0%", data: "95"})},
            {sheet: 1, row: 90, col: 1, json: {data: "10-6）"}},
            {sheet: 1, row: 90, col: 2, json: {data: "产品难度系数 K 值：K = 0.8-1.5 "}},
            {sheet: 1, row: 90, col: 3, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
            {sheet: 1, row: 91, col: 1, json: {data: ""}},
            {sheet: 1, row: 91, col: 2, json: {bgc: colorStepEnd, data: "粉体液体烤漆丝网印　小计： "}},
            {sheet: 1, row: 91, col: 3, json: styleSubTotal({data: "=C86*(C87+C88)*(1+(1-C89/100))*C90"})},
            {sheet: 1, row: 92, col: 1, json: {data: "* "}},
            {sheet: 1, row: 92, col: 2, json: {bgc: colorStep, data: "11)其它特殊处理： "}},
            {sheet: 1, row: 92, col: 3, json: {ta: "center", data: ""}},
            {sheet: 1, row: 93, col: 1, json: {data: "11-1）"}},
            {sheet: 1, row: 93, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 93, col: 3, json: styleInput({data: "12"})},
            {sheet: 1, row: 94, col: 1, json: {data: "11-2）"}},
            {sheet: 1, row: 94, col: 2, json: {data: "费用/只 "}},
            {sheet: 1, row: 94, col: 3, json: styleInput({fm: "money|¥|2|none", data: "0.34"})},
            {sheet: 1, row: 95, col: 1, json: {data: ""}},
            {sheet: 1, row: 95, col: 2, json: {bgc: colorStepEnd, data: "其它特殊处理　小计： "}},
            {sheet: 1, row: 95, col: 3, json: styleSubTotal({data: "=C94"})},
            {sheet: 1, row: 96, col: 1, json: {data: "* "}},
            {sheet: 1, row: 96, col: 2, json: {bgc: colorStep, data: "12)气密性测试 ： "}},
            {sheet: 1, row: 96, col: 3, json: {data: ""}},
            {sheet: 1, row: 97, col: 1, json: {data: "12-1）"}},
            {sheet: 1, row: 97, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 97, col: 3, json: styleInput({data: "56"})},
            {sheet: 1, row: 98, col: 1, json: {data: "12-2）"}},
            {sheet: 1, row: 98, col: 2, json: {data: "气测费/只 "}}, //气密性测试
            {sheet: 1, row: 98, col: 3, json: {fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=C97/3600*VLOOKUP('气密性测试',LOOKUP04!$A$2:$B$17,2,0)"}},
            {sheet: 1, row: 99, col: 1, json: {data: ""}},
            {sheet: 1, row: 99, col: 2, json: {bgc: colorStepEnd, data: "气密性测试　小计： "}},
            {sheet: 1, row: 99, col: 3, json: styleSubTotal({data: "=C98"})},
            {sheet: 1, row: 100, col: 1, json: {data: "* "}},
            {sheet: 1, row: 100, col: 2, json: {bgc: colorStep, data: "13)筛选包装 & 费率 /H： "}},
            {sheet: 1, row: 100, col: 3, json: {data: ""}},
            {sheet: 1, row: 101, col: 1, json: {data: "13-1）"}},
            {sheet: 1, row: 101, col: 2, json: {data: "工时（秒）/只 ： "}},
            {sheet: 1, row: 101, col: 3, json: styleInput({fm: "number", dfm: "0", data: "21"})},
            {sheet: 1, row: 102, col: 1, json: {data: "13-2）"}},
            {sheet: 1, row: 102, col: 2, json: {data: "人工费/只 "}},
            {sheet: 1, row: 102, col: 3, json: styleSubTotal({data: "=C101/3600*VLOOKUP('筛选和包装',LOOKUP04!$A$2:$B$17,2,0)"})},
            {sheet: 1, row: 103, col: 1, json: {data: "13-3）"}},
            {sheet: 1, row: 103, col: 2, json: {data: "包材费： "}},
            {sheet: 1, row: 103, col: 3, json: styleInput({fm: "money|¥|2|none", data: "0.35"})},
            {sheet: 1, row: 104, col: 1, json: {data: ""}},
            {sheet: 1, row: 104, col: 2, json: {bgc: colorStepEnd, data: "筛选包装　小计："}},
            {sheet: 1, row: 104, col: 3, json: styleSubTotal({data: "=C102+C103"})},
            {sheet: 1, row: 105, col: 1, json: {data: ""}},
            {sheet: 1, row: 105, col: 2, json: {bgc: colorSect, data: "各工序费用　合计："}},
            {sheet: 1, row: 105, col: 3, json: {bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true, data: "=C38+C48+C52+C59+C64+C69+C73+C77+C83+C91+C95+C99+C104"}},
            //{fm: "number", dfm: "0%",
            {sheet: 1, row: 106, col: 1, json: {data: "* "}},
            {sheet: 1, row: 106, col: 2, json: {data: "管销加利润%"}},
            {sheet: 1, row: 106, col: 3, json: {fm: "number", dfm: "0%", data: "20"}},
            {sheet: 1, row: 107, col: 1, json: {data: "* "}},
            {sheet: 1, row: 107, col: 2, json: {data: "管销加利润"}},
            {sheet: 1, row: 107, col: 3, json: styleSubTotal({data: "=C105*C106/100"})},
            {sheet: 1, row: 108, col: 1, json: {data: "* "}},
            {sheet: 1, row: 108, col: 2, json: {data: "其他五金配件： "}},
            {sheet: 1, row: 108, col: 3, json: styleInput({fm: "money|¥|2|none", data: "0"})},
            {sheet: 1, row: 109, col: 1, json: {data: "* "}},
            {sheet: 1, row: 109, col: 2, json: {data: "组装费： "}},
            {sheet: 1, row: 109, col: 3, json: styleInput({fm: "money|¥|2|none", data: "0"})},
            {sheet: 1, row: 110, col: 1, json: {data: ""}},
            {sheet: 1, row: 110, col: 2, json: {data: "组装 小计： "}},
            {sheet: 1, row: 110, col: 3, json: styleSubTotal({data: "=C108+C109"})},
            {sheet: 1, row: 111, col: 1, json: {data: ""}},
            {sheet: 1, row: 111, col: 2, json: {data: "合計（未稅）："}},
            {sheet: 1, row: 111, col: 3, json: styleSubTotal({data: "=C105+C107+C110"})},
//            {sheet: 1, row: 112, col: 0, json: {hidden: true}},
//            {sheet: 1, row: 113, col: 0, json: {hidden: true}},
//            {sheet: 1, row: 114, col: 1, json: {data: ""}},
//            {sheet: 1, row: 114, col: 2, json: {data: "总计（美金）： "}},
//            {sheet: 1, row: 114, col: 3, json: {fm: "money|$|2|none", dsd: "ed", cal: true, data: "=(C111/6.35)"}},
            {sheet: 1, row: 112, col: 1, json: {data: ""}},
            {sheet: 1, row: 112, col: 2, json: {data: "总计（美金）： "}},
            {sheet: 1, row: 112, col: 3, json: {fm: "money|$|2|none", dsd: "ed", cal: true, data: "=(C111/6.35)"}},
                {sheet: 1, row: 113, col: 1, json: {data: "生成", it: "button", btnStyle: "color: blue; font-weight: bold;", onBtnClickFn: "CUSTOM_BUTTON_CLICK___MAKE_EXCEL"}}
        //json: { data: "www.google.com", link: "www.google.com"} },
//                    {sheet: 1, row: 113, col: 2, json: { data: "www.google.com", link: "http://www.google.com"} },
  
        ]
    };
    // =============================================================================================
    // ok inject data now ...
//    var json_default = {
//		fileName: "Employee Directory",
//		sheets:[{id:1, name:"Main view", actived:true, color:"orange"}],
//		floatings: [
//	        { sheet:1, name:"colGroups", ftype:"colgroup", json: "[{level:1, span:[2,3]}, {level:1, span:[4,6]}]" },
//	    ],
//		cells:[
//		    { sheet: 1, row: 0, col: 0, json: { height: 20, va: "middle"} },
//		    { sheet: 1, row: 0, col: 1, json: { data: "ID", width: 50, dcfg: "{dt:0, io:true, min:0, max:10000, op:0, ignoreBlank: true, titleIcon: \"number\"}", ticon:"number" } },
//			{ sheet: 1, row: 0, col: 2, json: { data: "Name", width: 100, ticon:"profile"} },
//			{ sheet: 1, row: 0, col: 3, json: { data: "Dept.(Remote)", width: 130, drop: "list", dcfg: "{dt:15, url: \"fakeData/dropdownList\", titleIcon:  \"remoteList\"}", ticon:"remoteList" } },
//			{ sheet: 1, row: 0, col: 4, json: { data: "Email", width: 110, dcfg: "{dt:9}", ticon:"email" } },
//			{ sheet: 1, row: 0, col: 5, json: { data: "Phone", width: 100, dcfg: "{dt:8}", ticon:"phone" } },
//			{ sheet: 1, row: 0, col: 6, json: { data: "Gender", width: 80, drop: "list", dcfg: "{dt:13, list: [\"Male\",\"Female\"]}", ticon:"dropdown" } },
//			{ sheet: 1, row: 0, col: 7, json: { data: "Birth date", width: 120, drop: "date", fm: "date", dfm: "F d, Y", ticon:"calendar"  } },			
//			{ sheet: 1, row: 0, col: 8, json: { data: "Contact picker", width: 170, ticon:"contact", beforeEdit: "_beforeeditcell_" } },
//			{ sheet: 1, row: 0, col: 9, json: { data: "Manager?", width: 100, it: "checkbox", itchk: false, ta: "center", ticon:"checkbox" } },
//			{ sheet: 1, row: 0, col: 10, json: { data: "Images", width: 130, dcfg: "{dt:7}", ticon:"image" } },
//			{ sheet: 1, row: 0, col: 11, json: { data: "Salary", dcfg: "{dt:11, format: \"money|$|2|none|usd|true\"}",  ticon:"money_dollar" } },
//			{ sheet: 1, row: 0, col: 12, json: { data: "Percent", dcfg: "{dt:12, format: \"0.00%\"}",  ticon:"percent" } },
//			{ sheet: 1, row: 0, col: 13, json: { data: "Notes", dcfg: "{dt:14, titleIcon: \"textLong\"}",  ticon:"textLong" } },
//			
//			{ sheet: 1, row: 1, col: 1, json: { data: 1 } },
//			{ sheet: 1, row: 1, col: 2, json: { data: 'Jerry Marc' } },
//			{ sheet: 1, row: 1, col: 3, json: { render:'dropRender', data: 'HR Dept', dropId: 1} },
//			{ sheet: 1, row: 1, col: 4, json: { data: 'john.marc@abc.com'} },
//			{ sheet: 1, row: 1, col: 5, json: { data: '1 (888) 456-7654'} },
//			{ sheet: 1, row: 1, col: 6, json: { data: 'Female'} },
//			{ sheet: 1, row: 1, col: 7, json: { data: '1982-01-15', fm: "date", dfm: "F d, Y" } },
//			{ sheet: 1, row: 1, col: 8, json: { render:'contactRender', data: "Eva Mat, John Marc", itms: '[{name: "Eva Mat", email: "eva@gmail.com", id: 8}, {name: "John Marc", email: "john@abc.com", id: 9}]' } },
//			{ sheet: 1, row: 1, col: 10, json: { render:'attachRender', itms: '[{aid: "rT7KfpHA8cI_", url: "sheetAttach/downloadFile?attachId=rT7KfpHA8cI_", type: "img", name: "blue.jpg"},{aid: "2ZisVQ1-*Lo_", url: "sheetAttach/downloadFile?attachId=2ZisVQ1-*Lo_", type: "img", name: "green.jpg"}]' } },
//			{ sheet: 1, row: 1, col: 11, json: { data: 82334.5678 } },
//			{ sheet: 1, row: 1, col: 12, json: { data: 0.96 } },
//			{ sheet: 1, row: 1, col: 13, json: { data: 'This is notes, it is a long text. Double click to edit it.' } },
//			
//			{ sheet: 1, row: 2, col: 1, json: { data: 2 } },
//			{ sheet: 1, row: 2, col: 2, json: { data: 'Dave Smith' } },
//			{ sheet: 1, row: 2, col: 3, json: { render:'dropRender', data: 'Software Dept', dropId: 2} },
//			{ sheet: 1, row: 2, col: 4, json: { data: 'dave.smith@abc.com'} },
//			{ sheet: 1, row: 2, col: 5, json: { data: '1 (888) 231-7654'} },
//			{ sheet: 1, row: 2, col: 6, json: { data: 'Male'} },
//			{ sheet: 1, row: 2, col: 7, json: { data: '1980-01-15', fm: "date", dfm: "F d, Y" } },
//			{ sheet: 1, row: 2, col: 8, json: { render:'contactRender', data: "Christina Angela, Marina Chris", itms: '[{name: "Christina Angela", email: "christina@gmail.com", id: 4}, {name: "Marina Chris", email: "marina@abc.com", id: 6}]' } },
//			{ sheet: 1, row: 2, col: 10, json: { render:'attachRender', itms: '[{aid: "CIBHu3ffG8Q_", url: "sheetAttach/downloadFile?attachId=CIBHu3ffG8Q_", type: "img", name: "admin.png"},{aid: "VcrhEYAyrzA_", url: "sheetAttach/downloadFile?attachId=VcrhEYAyrzA_", type: "img", name: "asset.png"}]' } },
//			{ sheet: 1, row: 2, col: 11, json: { data: 81234.5678 } },
//			{ sheet: 1, row: 2, col: 12, json: { data: 0.95 } },
//			{ sheet: 1, row: 2, col: 13, json: { data: 'This is notes, it is a long text. Double click to edit it.' } },
//			
//			{ sheet: 1, row: 3, col: 1, json: { data: 3 } },
//			{ sheet: 1, row: 3, col: 2, json: { data: 'Kevin Featherstone' } },
//			{ sheet: 1, row: 3, col: 3, json: { render:'dropRender', data: 'Software Dept', dropId: 2} },
//			{ sheet: 1, row: 3, col: 4, json: { data: 'kevin@abc.com'} },
//			{ sheet: 1, row: 3, col: 5, json: { data: '1 (888) 232-7654'} },
//			{ sheet: 1, row: 3, col: 6, json: { data: 'Male'} },
//			{ sheet: 1, row: 3, col: 7, json: { data: '1990-01-15', fm: "date", dfm: "F d, Y" } },
//			{ sheet: 1, row: 3, col: 8, json: { render:'contactRender', data: "Christina Angela, Marina Chris", itms: '[{name: "Christina Angela", email: "christina@gmail.com", id: 4}, {name: "Marina Chris", email: "marina@abc.com", id: 6}]' } },
//			{ sheet: 1, row: 3, col: 10, json: { render:'attachRender', itms: '[{aid: "CIBHu3ffG8Q_", url: "sheetAttach/downloadFile?attachId=CIBHu3ffG8Q_", type: "img", name: "admin.png"},{aid: "VcrhEYAyrzA_", url: "sheetAttach/downloadFile?attachId=VcrhEYAyrzA_", type: "img", name: "asset.png"}]' } },
//			{ sheet: 1, row: 3, col: 11, json: { data: 81934.5678 } },
//			{ sheet: 1, row: 3, col: 12, json: { data: 0.98 } },
//			{ sheet: 1, row: 3, col: 13, json: { data: 'This is notes, it is a long text. Double click to edit it.' } }
//		]
//	};	

    //SHEET_API.loadData(SHEET_API_HD, json, null, this);
    SHEET_API.loadData(SHEET_API_HD, json, function () {
        SHEET_API.setMaxRowNumber(114);
//        SHEET_API.setMaxColNumber(8);//col H is NOT working well when copy/paste col
        SHEET_API.setMaxColNumber(9);
        
    }, this);

    //DOING TESTING SAVE DATA
    // SHEET_API.saveData(SHEET_API_HD, null, this);
    //  TESING 05/07 12:03
    var cellData = SHEET_API.getCell(SHEET_API_HD, 1, 100, 2);

// chck whether it is formula
//    if (cellData.cal) {
//        var formula = cellData.arg;
//        var value = cellData.value;
//        var result = cellData.data;
//        var format = cellData.fm;
//        var detailFormat = cellData.dfm;
//    } else {
//        var result = cellData.data;
//        var format = cellData.fm;
//        var detailFormat = cellData.dfm;
//    }
    console.log("=== TESTING getCell 1,100,2 =>13)筛选包装 & 费率 /H：");
    console.log(cellData.data);

    function twcloudExportCells() {
        var cellDataB, cellDataC, cellDataD;
        console.log("--- DOING twcloudExportCells --- ");
        for (var i = 0; i < 114; i++) {
            cellDataB = SHEET_API.getCell(SHEET_API_HD, 1, i, 2);
            cellDataC = SHEET_API.getCell(SHEET_API_HD, 1, i, 3);
            cellDataD = SHEET_API.getCell(SHEET_API_HD, 1, i, 4);

//            if (cellData.cal) {
//                var formula = cellData.arg;
//                var value = cellData.value;
//                var result = cellData.data;
//                var format = cellData.fm;
//                var detailFormat = cellData.dfm;
//            } else {
//                var result = cellData.data;
//                var format = cellData.fm;
//                var detailFormat = cellData.dfm;
//            }
            console.log(i + " " + cellDataB.data);
            console.log(i + " " + cellDataC.data);
            console.log(i + " " + cellDataD.data);

        }
    }
    twcloudExportCells();

    SHEET_API.setFocus(SHEET_API_HD, 2, 1);
    // add event listener - this shows the code to add customer function 
    var sheet = SHEET_API_HD.sheet;
    var editor = sheet.getEditor();
    editor.on('quit', function (editor, sheetId, row, col) {
        if (col === 1) {
            // this is the method to query customer existing backend and auto fill data
            // 
            // NOTE, GOOD EXAMPLE, MAYBE IMPLEMENT IT LATER 05/06 15:32
//            var employeeId = SHEET_API.getCellValue(SHEET_API_HD, sheetId, row, col).data;
//            if (employeeId)
//                AUTO_FILL_CUSTOMER_DATA_BY_EMPLOYEEID(employeeId, sheetId, row, col);
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
	