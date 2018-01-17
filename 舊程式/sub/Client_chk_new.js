//自行定義函數
//先判斷瀏覽器是否支援，若是不支援就自行定義

//定義trim
//由於trim此方法是在ECMAScript 第五版才引進。所以為了避免舊版的瀏覽器無法執行，可以自行定義該方法。
if (!String.prototype.trim) {
    String.prototype.trim = function () {
        return this.replace(/^\s+|\s+$/g, '');
    };
}

//定義execScript
//新版本的chrome已經不支援
if (!window.execScript) {
    //window.execScript = function execscript(text) {
    //    if (!typeof text === "string") return;
    //    //window.eval(text)
    //    new Function("", text)()
    //}

    window.execScript = function(text) {
        if (!text) return;

        // 構建script元素
        var script = document.createElement("SCRIPT");
        script.setAttribute("type", "text/javascript");
        script.text = text;

        var head = document.head || document.getElementsByTagName("head")[0]

        head.appendChild(script); // 開始執行
        head.removeChild(script); // 執行完清除
    }
}


//1.檢查欄位值函數
//2.數值函數
//3.日期、時間函數
//4.編碼函數
//5.人事/薪號函數
//6.其他通用函數

//----------檢查欄位值函數begin----------

//檢查單一欄位不可為空白
function chkNull(pFieldName, pobject) {
    if (pobject.value.trim() == "") {
        alert(pFieldName + "必須輸入!!!");
        pobject.focus();
        return true;
    }
    
    return false;
}

//檢查欄位長度
function chkLength(pobject) {
    /*
    var maxln = parseInt(obj.maxlength);
    var txt = obj.value;
    if (txt.length > maxln) {
        alert("长度超过 " + maxln + " 字元，将会截掉多出的字元 !");
        obj.value = txt.substr(0, maxln);
        obj.focus();
    }
    return true;
    */
    
    for (var x in pobject.attributes) {
        //alert(x + "-" + pobject.getAttribute(x));
        // 判斷該欄位是否有設定maxlength屬性
        if (x.toLowerCase() == "maxlength") {
            if (pobject.getAttribute(x) != "2147483647") {// maxlength的預設值，抓到此值表示該欄位並沒有設定maxlength
                if (pobject.value != "") {
                    var tStr1;
                    var tLen = 0;
                    var tCod;

                    for (var i = 0; i < pobject.value.length; i++) {
                        tStr1 = pobject.value.substr(i, 1);
                        //alert(tStr1);
                        tCod = tStr1.charCodeAt(0);
                        if (tCod >= 128 || tCod < 0)
                            tLen += 2;
                        else
                            tLen += 1;
                    }

                    if (tLen > pobject.maxLength) {
                        alert("長度過長，請檢查！");
                        pobject.focus();
                        return true;
                    }

                    return false;
                }
            }
        }
    }
}

// 檢查radio是否有值
function chkRadio(pobject, pmsg) {
    var sret = false;
    
    for (var i = 0; i < pobject.length; i++) {
        if (pobject[i].checked == true) {
            sret = true;
            break;
        }
    }
    
    if (sret == false) {
        alert("請選擇" + pmsg + "！");
    }
    
    return sret;
}

//檢查checkbox是否有值
function chkChkbox(pobject, pmsg) {
    var sret = false;

    for (var i = 0; i < pobject.length; i++) {
        if (pobject[i].checked == true) {
            sret = true;
            break;
        }
    }

    if (sret == false) {
        alert("請選擇" + pmsg + "！");
    }

    return sret;
}

//----------檢查欄位值函數end----------

//----------數值函數begin----------

//檢查欄位必須為數值
/*
function chkNum(pnm, pva) {
    nm = pnm.split("|");
    va = pva.split("|");
        
    for (z = 0; z < nm.length; z++) {
        if (va[z] != "") {
            if (isNaN(va[z])) {
                alert(nm[z] + "必須為數值!!!");
                return true;
            }
        }
        }
    return false;
}
*/
//檢查欄位必須為數值
function chkNum(pobject) {
    if (pobject.value.trim() != "") {
        if (isNaN(pobject.value)) {
            alert("此欄位必須輸入數值!!!");
            pobject.focus();
            return true;
        }
    }

    return false;
}

//檢查單一欄位不可為小數
function chkInt(pobject) {
    var tvalue;
    var tint;

    if (pobject.value.trim() != "") {
        var intrep = pobject.value.trim().match(/^[0-9]+$/);    //判斷是否為正整數
        var nintrep = pobject.value.trim().match(/^[-]{1}[0-9]+$/);     //判斷是否為負整數
        if (intrep == null && nintrep == null) {
            alert("此欄位必須為整數，請重新輸入!!!");
            pobject.focus();
            return true;
        }
    }

    return false;
}

function FormatNumber(pnum, ppos) {
    var sret;

    if (!isNaN(pnum)) {
        var num = new Number(pnum);

        sret = num.toFixed(ppos);
    }
    
    return sret;
}

//----------數值函數end----------

//----------日期、時間函數begin----------

//檢查日期格式
function chkdateformat(pobject) {
    if (pobject.value.trim() == "") return true;

    //JavaScript substr 與 substring 的差異
    //String.substr( Start , Length ) 
    //String.substring( Start , End )
    //substr 的第二個參數是字串長度，可自由設定，如果沒有填寫，則自動取至 String 字串的最後一個字符
    //substring 的第二個參數是結尾字符，自動擷取至該字符的前一個字符，如果沒有填寫，一樣擷取至最後一個字符。	
    if (pobject.value.length == 8 && isNaN(pobject.value) == false)
        pobject.value = pobject.value.substring(0, 4) + "/" + pobject.value.substring(4, 6) + "/" + pobject.value.substring(6, 8);

    if (IsDate(pobject.value.trim()) == false) {  //javascript無IsDate函數，需自行處理
        alert("日期格式錯誤，請重新輸入!!! 日期格式:YYYY/MM/DD");
        pobject.focus();
        return true;
    }
    else {
        //pobject.value = cdate(pobject.value)
        pobject.value = pobject.value.replace("-", "/");
        dateObj = new Date(pobject.value);
        pobject.value = dateObj.getFullYear() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getDate();
    }

    return false;
}

//檢查日期格式
function IsDate(pobject) {
    var rep = new RegExp("^([0-9]{4})[./-]{1}([0-9]{1,2})[./-]{1}([0-9]{1,2})$");
    var strDateValue;
    var sret = true;

    if ((strDateValue = rep.exec(pobject)) != null) {
        var i;
        i = parseFloat(strDateValue[1]);
        if (i <= 0 || i > 9999) { //年
            sret = false;
        }
        i = parseFloat(strDateValue[2]);
        if (i <= 0 || i > 12) { //月
            sret = false;
        }
        i = parseFloat(strDateValue[3]);
        if (i <= 0 || i > 31) { //日
            sret = false;
        }
    }
    else {
        sret = false;
    }

    return sret;
}

// 檢查年月格式
function chkYM(pFieldName, pobject) {
    if (pobject.value.trim() == "") return true;

    if (pobject.value.length != 6) {
        alert(pFieldName + "年月格式錯誤，請重新輸入!!! 年月格式:YYYYMM");
        pobject.focus();
        return true;
    }

    if (IsDate(pobject.value.substring(0, 4) + "/" + pobject.value.substring(4, 6) + "/1") == false) {
        alert(pFieldName + "年月格式錯誤，請重新輸入!!! 年月格式:YYYYMM");
        pobject.focus();
        return true;
    }

    return false;
}

// 檢查時間格式(HH:MM)
function chktime(pmsg, ptime) {
	var sret;
	
	sret = true;
	
	if (IsTimeValue(ptime) == false) {
	    alert(pmsg +"輸入錯誤!!!");
	    sret = false;
	}
	
	return sret;
}

function DateAdd(timeU, byMany, dateObj) {
    var sret;

    var millisecond = 1;
    var second = millisecond * 1000;
    var minute = second * 60;
    var hour = minute * 60;
    var day = hour * 24;
    var year = day * 365;

    var newDate;
    var dVal = new Date(dateObj).valueOf();
    var dObj = new Date(dateObj);

    switch (timeU) {
        case "ms":
            newDate = new Date(dVal + millisecond * byMany);
            break;
        case "s":
            newDate = new Date(dVal + second * byMany);
            break;
        case "mi":
            newDate = new Date(dVal + minute * byMany);
            break;
        case "h":
            newDate = new Date(dVal + hour * byMany);
            break;
        case "d":
            newDate = new Date(dVal + day * byMany);
            break;
        case "m":
            newDate = new Date(dObj.setMonth(dObj.getMonth() + byMany));
            break;
        case "y":
            newDate = new Date(dVal + year * byMany);
            break;
    }

    sret = newDate.getFullYear() + "/" + (newDate.getMonth() + 1) + "/" + newDate.getDate();

    return sret;
}

function DateDiff(interval, pDate1, pDate2) {
    var sret;
    var objDate1;
    var objDate2;

    objDate1 = new Date(pDate1);
    objDate2 = new Date(pDate2);

    //若參數不足或 objDate 不是日期物件則回傳 undefined 
    if (arguments.length < 3 || objDate1.constructor != Date || objDate2.constructor != Date) return undefined;

    switch (interval) {
        case "s": //計算秒差
            sret = parseInt((objDate2 - objDate1) / 1000);
            break;
        case "n":   //計算分差
            sret = parseInt((objDate2 - objDate1) / 60000);
            break;
        case "h":   //計算時差
            sret = parseInt((objDate2 - objDate1) / 3600000);
            break;
        case "d":   //計算日差
            sret = parseInt((objDate2 - objDate1) / 86400000);
            break;
        case "w":  //計算週差
            sret = parseInt((objDate2 - objDate1) / (86400000 * 7));
            break;
        case "m":  //計算月差
            sret = (objDate2.getMonth() + 1) + ((objDate2.getFullYear() - objDate1.getFullYear()) * 12) - (objDate1.getMonth() + 1);
            break;
        case "y":   //計算年差
            sret = objDate2.getFullYear() - objDate1.getFullYear();
            break;
        default:    //輸入有誤
            sret = undefined;
            break;
    }

    return sret;
} 

//取得日期中的年月日(去除時分秒)
function GetDateString(pdate) {
    var sret = "";

    var dateStr;

    if (pdate.trim() != "") {
        dateStr = pdate;
        
        if (dateStr.indexOf("上午") != -1)
            dateStr = dateStr.replace("上午", "");

        if (dateStr.indexOf("下午") != -1)
            dateStr = dateStr.replace("下午", "");    
            
        var dateObj = new Date(dateStr);
        //getMonth取得月份 {[一月] 0 - [十二月] 11}，故需再加1
        sret = dateObj.getFullYear() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getDate();   
    }
    
    //alert(sret);
    return sret;
}

//取得日期格式YYYY/MM/DD(有補0)
function cformatdate(pdate) {
    var sret = "";

    var dateStr = pdate;
	
	if (pdate.trim() != "") {
	    if (dateStr.indexOf("上午") != -1)
	        dateStr = dateStr.replace("上午", "");

	    if (dateStr.indexOf("下午") != -1)
	        dateStr = dateStr.replace("下午", "");    	  
	          
	    var dateObj = new Date(dateStr);
	    //getMonth取得月份 {[一月] 0 - [十二月] 11}，故需再加1
        //JavaScript slice 基本語法
	    //String.slice( Start , End )
        //slice 可以從字串中擷取某一段字串出來，用法與 JavaScript 的substring類似，但slice 比較特別的地方在於可以從字串尾端開始計算位置
        //語法中的開頭 String 是原始字串，slice 函式小括號內的＂Start＂與＂End＂用以標示擷取的區間，兩者均可為負數
        //Start 如果是負數，則表示從字串的最尾處開始，-1 代表最後一個字，-2 代表倒數第二個字，-3 代表倒數第三個字，以此類推
        //End 的概念也是一樣
        //End 如果未填寫，則代表 slice 函式從字串的第 Start 的字，開始擷取到字串的最後一個字。	    
	    sret = dateObj.getFullYear() + "/" + ("0" + (dateObj.getMonth() + 1)).slice(-2) + "/" + ("0" + dateObj.getDate()).slice(-2);
	}

	//alert(sret);
    return sret;
}

//取得今天日期
function GetToday() {
    var sret;
    
    var dateObj = new Date();

    sret = dateObj.getFullYear() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getDate();  

    //alert(sret);
    return sret; 
}

//點選日曆上之日期之後，自動focus在該text欄位
function mydate(ddts, sObj) {
    //alert(ddts[0]);
    $(sObj).focus();
}

//for動態新增(多筆)的日期欄位，設定日期屬性
function SetDatePick(did) {
    $("input:text[name='" + did + "']").datepick({ onSelect: function(dates) { mydate(dates, this); } });
}

//for動態新增(多筆)的日期欄位，解除日期屬性
function DisDatePick(did) {
    $("input:text[name='" + did + "']").datepick("disable");
}

// 檢查是否為時間格式
function IsTimeValue(ptime) {
    var sret = true;

    if (ptime.trim() == "") return true;

    // 檢查必須是正確的時間格式
    // 1. 必須為5碼字串
    // 2. 前兩碼為數字，範圍：00~23
    // 3. 第三碼為冒號(:)
    // 4. 後兩碼為數字，範圍為：00~59

    if (ptime.length != 5) {
        sret = false;
    }
    else if (isNaN(ptime.substring(0, 2))) {
        sret = false;
    }
    else if (!(parseInt(ptime.substring(0, 2)) >= 0 && parseInt(ptime.substring(0, 2)) <= 23)) {
        sret = false;
    }
    else if (ptime.substring(2, 3) != ":") {
        sret = false;
    }
    else if (isNaN(ptime.substring(3, 5))) {
        sret = false;
    }
    else if (!(parseInt(ptime.substring(3, 5)) >= 0 && parseInt(ptime.substring(3, 5)) <= 59)) {
        sret = false;
    }

    //alert(sret);
    return sret;
}

//轉換為有時分秒之日期格式
function TimeValue(ptime) {
    var sret = "";
    
    if (ptime != "") {
        sret = new Date("2000/1/1" + " " + ptime).valueOf();
    }

    //alert(sret);
    return sret;
}

//----------日期、時間函數end----------

//----------編碼函數begin----------

//URLENCODE-Big5
function URLEncodeBig5(pstr) {
    var sret = "";
    
    /*
    var strSpecial = "!\"#$%&'()*+,/:;<=>?[]^`{|}~%";
    var tt = "";
    for (var i = 0; i < pstr.length; i++) {
        var chr = pstr.charAt(i);
        var c = str2asc(chr);
        tt += chr + ":" + c + "n";
        if (parseInt("0x" + c) > 0x7f) {
            sret += "%" + c.slice(0, 2) + "%" + c.slice(-2);
        } else {
            if (chr == " ")
                sret += "+";
            else if (strSpecial.indexOf(chr) != -1)
                sret += "%" + c.toString(16);
            else
                sret += chr;
        }
    } 
    */

    //escape、encodeURI、encodeURIComponent三種方法的差異：
    //1.英文字、數字、-、_、.、* 這些字不管是哪一種Javascript URLEncode方式，都是不會被encode的
    //2.各Encode方法，都是Encode為Unicode，但escape是Encode為UTF-16、而encodeURI與encodeURIComponent則是UTF-8
    //3.因為 Javascript 都是 Encode 為 Unicode ，因此如有特殊用途須使用Big5編碼，需自行處理，例如 URI 中的 MailTo:

    //URL有最大長度的限制(2083個字元)，但實際測試的結果不足2083個字元
    //有些字元可能在產生mailto的url時，會再被編碼成HtmlEncode，例如&字元，會被HtmlEncode為&amp;
    //所以在最後呈現在Html上的字數限制計算上，要十分注意。不然會有無法開啟mailto、或有錯誤訊息的狀況發生
    
    var url = "../xml/xml_urlencode_big5.aspx";
    var data = "SearchStr=" + encodeURIComponent(pstr);
    //window.open(url + "?" + data);
    $.ajax({
        url: url, data: data, async: false, dataType: "xml",
        success: function(xmldoc) {
            if ($(xmldoc).find("Found").text() == "Y") {
                sret = $(xmldoc).find("urlencodebig5").text();
                //alert("javascript:" + sret);
                //window.open("mailto:?subject=" + sret);
            }
        }, error: function() { alert("URLEncodeBig5求取資料時錯誤 !"); }           
    });
     
    //alert(sret);
    return sret;
}

function encodeToHex(str) {
    var r="";
    var e=str.length;
    var c=0;
    var h;
    while(c<e){
        h=str.charCodeAt(c++).toString(16);
        while(h.length<3) h="0"+h;
        r+=h;
    }
    return r;
}

function str2asc(str) {
    return str.charCodeAt(0).toString(16);
}

function asc2str(str) {
    return String.fromCharCode(str);
}

function ToUnicode(str) {
    var sret = "";

    sret = str;

    return sret;
}

//----------編碼函數end----------

//----------人事/薪號函數begin----------

// 求取該單位所屬部門
function get_grpid_below(grpclass, grpid) {
    var sret = "";

    var url = "../xml/xml_grpid_below.aspx";
    var data = "grpclass=" + grpclass + "&grpid=" + grpid;
    //window.open(url + "?" + data);
    $.ajax({
        url: url, data: data, async: false, dataType: "xml",
        success: function(xmldoc) {
            if ($(xmldoc).find("Found").text() == "Y") {
                sret = $(xmldoc).find("grpid").text().trim();
            }
        }, error: function() { alert("get_grpid_below求取資料時錯誤!"); }
    });

    return sret;
}

//轉換人事主檔薪號
function GetPersonScode(pscode) {
    var sret = "";

    if (pscode.trim() == "") return sret;

    if (pscode.trim().toLowerCase() == "admin")
        sret = "admin";
    else {
        sret = pscode.trim().substring(1);
        sret = ("0000" + sret).slice(-4);
    }

    //alert(sret);
    return sret;
}

//轉換人事主檔薪號
function TranPersonScode(pscode) {
    var sret = "";

    if (pscode.trim() == "") return sret;

    //若薪號第一碼不是數字，則判斷scode檔有無此薪號
    if (isNaN(pscode.trim().substring(0, 1))) {
        var sql = "SELECT scode, sc_name FROM scode WHERE scode = '" + pscode.trim() + "'";

        var url = "../xml/XmlGetSqlDataSysctrl.aspx";
        var data = "SearchSql=" + sql;
        //window.open(url + "?" + data);
        $.ajax({
            url: url, data: data, async: false, dataType: "xml",
            success: function(xmldoc) {
                if ($(xmldoc).find("Found").text() == "Y") {
                    sret = GetPersonScode(pscode);
                }
                else {
                    alert(pscode.trim() + "此薪號不存在！");
                    //pscode.focus();
                }
            }, error: function() { alert("TranPersonScode求取資料時錯誤 !"); }
        });
    }

    //alert(sret);
    return sret;
}

//----------人事/薪號函數end----------

//----------其他通用函數begin----------

//頁籤
function settab(k) {
    $("#CTab td.notab").removeClass("notab").addClass("seltab");
    $($("#CTab td.tab")[k]).removeClass("seltab").addClass("notab");
    $("#Cont div.tabCont").hide();
    $($("#Cont div.tabCont")[k]).show();
    return true;
}

// 解除所有欄位的disabled屬性
function openfield(form) {
    var x = document.forms[form];

    for (var i = 0; i < x.elements.length; i++) {
        if (x.elements[i].tagname != "fieldset") {
            //filedset沒有type此屬性
            switch (x.elements[i].type) {
                case "select-one":
                case "textarea":
                case "radio":
                case "checkbox":
                case "text":
                    x.elements[i].disabled = false;
                    break;
            }
        }
    }
    
}

function chkTest_onclick() {
	if (reg.chkTest.checked == true) 
		document.getElementById("ActFrame").style.display = "";
    else 
        document.getElementById("ActFrame").style.display = "none";
}

function IIf(bBool, trueStr, falseStr) {
	var sret;

	if (bBool)
	    sret = trueStr;
	else
	    sret = falseStr;

    return sret;
}

//----------其他通用函數end----------