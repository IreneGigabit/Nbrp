<script Language=javaScript>
function archkNaN(pnm,pva) //有輸入必須為數值
{
	nm = pnm.split("|");
	va = pva.split("|");
	for (z=0;z<nm.length;z++) {
		if (va[z]!="")
		{	if (isNaN(va[z]))
			{	alert(nm[z]+"必須為數值!!!");
				return true;
			}
		}
	}
	return false;
}
function ChkNumValue(pNumStr,pChkZero,pChkInt) //編審數值欄位的輸入值
{   
/*	傳入參數:	1.pNumStr:數值
				2.pChkZero=0:不Check欄位值正負數,1:Check欄位值必須大於0,2:Check欄位值必須>=0
				3.pChkInt=0:不Check欄位值為整數,1:Check欄位值為整數
	傳回參數:	Error=true, Success=false	
*/
	if (pNumStr == "") {return false;}
	if (isNaN(pNumStr) == true) {
		alert("欄位值須為數值 !");
		return true;
	}
	switch(pChkZero)
	{
		case 0:
			break;
		case 1:
			if (pNumStr <= 0) {
				alert("數值必須大於 0 !");
				return true;
			}
			break;
		case 2:
			if (pNumStr < 0) {
				alert("數值必須大於等於 0 !");
				return true;
			}
			break;
	}
    if (pChkInt == 1) {
		if (parseInt(pNumStr) != pNumStr) {
			alert("欄位值須為整數 !");
			return true;
		}
    }
    return false;
}
function ChkNumFormat(pNumStr,pDefineInt,pPoint,pType)
{
/*功能說明:	pType=0:四捨五入到小數第pPoint位,1:無條件捨去到小數第pPoint位,
					2:無條件進位到小數第pPoint位
	傳入參數:	1.pNumStr:數值,
				2.pDefineInt:Table Layout所訂該欄位整數部分長度(1,2....)
				3.pPoint:Table Layout所訂該欄位小數部分長度, 亦即計算到小數第N位(0,1,2....),
				4.pType:0:四捨五入,1:無條件捨去,2:無條件進位
	傳回參數:Error:Return Array(true,""),No Error: Return Array(false,轉換正確的數字)
*/
	var tPointNum,tTempNum,tResult;
	var tPos,ixI;
	var swFmtErr;

	if (pNumStr == "") {return Array(false,pNumStr);}
	swFmtErr=false;
	if (isNaN(pNumStr) == true) //確定輸入值是數字
	{   alert("輸入值不是數字 !");
		return Array(true,"");
	}
	//確定輸入值格式正確(ex:輸入值的整數位數符合Table Define)
	tPos = pNumStr.indexOf(".");
	if (tPos > 0)
	{	tIntNumber=pNumStr.substring(0,tPos);
		tPointNumber = pNumStr.substring(tPos+1,pNumStr.length);
		if (tIntNumber.length > pDefineInt)	{swFmtErr=true;}
	}
	else
	{   if (pNumStr.length > pDefineInt) {swFmtErr=true;}}
	if (swFmtErr==true)
	{   alert("輸入格式錯誤 !");
		return Array(true,"");
	}

	tPointNum=1;

	//到小數點n位就有n個0
	for(var ixI=0;ixI<pPoint;ixI++) {tPointNum = tPointNum * 10;}
	tTempNum = (pNumStr * tPointNum);
	switch(pType)
	{
		case 0: //四捨五入
			tTempNum = Math.round(tTempNum);
			break;
		case 1: //無條件捨去
			tTempNum = parseInt(tTempNum);
			break;
		case 2: //無條件進位
			tTempNum = parseInt(tTempNum) + 1;
			break;
	}
	if (tPos > 0) //若是有小數判斷小數位數是否符合判斷式
	{
		if (tPointNumber.length < pPoint)
		{   tResult = pNumStr;}
		else
		{   tResult = (tTempNum/tPointNum);}
	}
	else
	{   tResult=pNumStr;}  //若是整數傳整數回去
	return Array(false,tResult);
}
</script>
