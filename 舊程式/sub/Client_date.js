<script Language="javascript">
/*	chkSEDate: 起始日不可大於迄止日
	chkEDate:	檢查西元年輸入正確否
*/
function chkSEDate(pSdate,pEdate) //起始日不可大於迄止日
{
	if (pSdate=="" && pEdate=="")
	{	return true;}
	if (ChkEDate(pSdate)[1] > ChkEDate(pEdate)[1])
	{	alert("起始日不可大於迄止日");
		return true;}
	return false;
}
function ChkEDate(pDateStr)  //編審西元年
{
//傳入參數:YYYYMMDD(ex:19990101) or YYYY/MM/DD(ex:1999/01/01 or 1999/1/1)
//傳回參數:Error:Return Array(true,""), Success:Return Array(false,標準格式西元年)
	if (pDateStr == "")
	{   return Array(false,pDateStr);}
	tPos = pDateStr.indexOf("/");
	if (tPos == -1)   //輸入19990101 (輸入的資料若未含/,須輸入8位)
	{
		if (pDateStr.length != 8) 
		{
			alert(pDateStr+"日期格式錯誤 (ex:YYYYMMDD) !");
			return Array(true,"");
		}
		else
		{
			pDateStr.substring(0,8);
			tYYYY = parseInt(pDateStr.substring(0,4),10);
			tMM = pDateStr.substring(4,6);
			tMM1 = parseInt(tMM,10) - 1;
			tDD = pDateStr.substring(6,8);
		}
	}
	else         //輸入1999/01/01 or 輸入1999/1/1
	{
		tYYYY = parseInt(pDateStr.substring(0,tPos),10) //年
		if (tYYYY.toString().length != 4)
		{
			alert("日期格式錯誤 (ex:YYYY/MM/DD) !");
			return Array(true,"");
		}
		tDateStr = pDateStr.substring(tPos+1);
		tPos   = tDateStr.indexOf("/");
		tMM    = tDateStr.substring(0,tPos);      //月
		if (tMM.length == 1) {tMM = "0" + tMM;}
		tMM1   = parseInt(tMM,10) - 1;
		tDD    = tDateStr.substring(tPos+1);      //日
		if (tDD.length == 1) {tDD = "0" + tDD;}
	}
	if (tYYYY>=0 && tYYYY<=9999) {tTYYYY=tYYYY;}
	if (tMM>=1 && tMM<=12) {tTMM=tMM;}
	if (tDD>=1 && tDD<=31) {tTDD=tDD;}
//	var tDate = new Date(tYYYY,tMM1,tDD).toLocaleString();    //將日期轉為MM/DD/YYYY
	//需配合server時間設定
//	tTYYYY = tDate.substring(6,10);
//	tTMM = tDate.substring(0,2);
//	tTDD = tDate.substring(3,5);
	if (tYYYY != tTYYYY || tMM != tTMM)
	{
		alert(pDateStr+"輸入無效的西元日期 !");
		return Array(true,"");
	}
	else
	{	tResult=tTYYYY+"/"+tMM+"/"+tDD;}
	return Array(false,tResult);
}
</script>