<script Language=javaScript>
function ChkID(pUserID,pType) //編審(身份證字號,統一編號)
{  
/*	傳入參數:	1.pUserID:身份證字號:10位文數字,統一編號:8位數字
				2.pType:0:編審身份證字號,1:編審統一編號
	傳回參數:Boolean:True不正確,False正確	*/
    var ix_I;
    if (pUserID == "") {return false;}
    switch(pType)
    {
		case 0: //編審身份證字號
			var tAreaNo;
			var tCheckSum;
			var tAreaCode;
			var tSecondID;         //身份證第二碼

			pUserID = pUserID.toUpperCase();
			tAreaCode = pUserID.substr(0,1);
			if (pUserID.length != 10)  //確定身份證字號有10碼
			{	alert("輸入無效的身份證字號 (ex:資料長度錯誤) !");
				return true;
			}
			if (tAreaCode.valueOf()<"A" && tAreaCode.valueOf()>"Z")  //確定首碼在A-Z之間
			{	alert("輸入無效的身份證字號 (ex:首碼應介於A-Z之間) !");
				return true;
			}
			if (isNaN(parseInt(pUserID.substring(1,10),10)) == true)  //確定2-10碼是數字
			{	alert("輸入無效的身份證字號 (ex:第2-10碼須是數字) !");
				return true;
			}
			//身份證號碼第 2 碼必須為 1 或 2
			tSecondID = pUserID.substr(1,1);
			if (tSecondID != 1 && tSecondID != 2) 
			{    alert("輸入無效的身份證字號 !");
			    return true;
			}
			//取得首碼對應的區域碼，A ->10, B->11, ..H->17,I->34, J->18...
			tAreaNo = "ABCDEFGHJKLMNPQRSTUVXYWZIO".search(tAreaCode) + 10;
			pUserID = tAreaNo.toString(10) + pUserID.substring(1,10);   

			//  取得CheckSum的值,核對身份證號碼是否正確
			//  A = 第1碼, A0 = 第1碼*(10-1), A1 = 第2碼*(10-2), A2 = 第3碼*(10-3)
			//  A3 = 第4碼*(10-4), A4 = 第5碼*(10-5), A5 = 第6碼*(10-6)
			//  A6 = 第7碼*(10-7), A7 = 第8碼*(10-8), A8 = 第9碼*(10-9)
			//  CheckSum = A+A0+A1+A2+A3+A4+A5+A6+A7+A8

			tCheckSum = parseInt(pUserID.substr(0,1),10) + parseInt(pUserID.substr(10,1),10);
			for(ixI=1;ixI<=9;ixI++)
			{	tCheckSum = tCheckSum + parseInt(pUserID.substr(ixI,1),10)*(10-ixI);}
			if ((tCheckSum % 10) != 0)
			{    alert("輸入無效的身份證字號 !");
			     return true;
			}
			return false;
			break;

		case 1: //編審統一編號
			var tSum=0;
			var tDiv=0;
			var tMod=0;
			var tStr="12121241";
			         
			if (parseInt(pUserID.substring(0,8),10)!=pUserID) //確定1-8碼是數字 
			{   alert("輸入無效的統一編號 (ex:須為8位數字)!");
				return true;
			} 
			if (isNaN(parseInt(pUserID.substring(0,8),10)) == true) //確定1-8碼是數字
			{   alert("輸入無效的統一編號 (ex:須為8位數字)!");
				return true;
			} 
			for(ixI=0; ixI<=7; ixI++)//套公式編審
			{	tDiv=parseInt(parseInt(pUserID.substr(ixI,1),10)*parseInt(tStr.substr(ixI,1))/10);
				tMod=parseInt(parseInt(pUserID.substr(ixI,1),10)*parseInt(tStr.substr(ixI,1))%10);
				tSum=tSum+tDiv+tMod;
			}
			tSum=parseInt(tSum%10);
			         
			if ((tSum==0 || tSum==9) && pUserID.substr(6,1)=="7")
			{    return false;} //正確
			if (tSum==0)
			{   return false;} //正確
			else
			{	alert("輸入無效的統一編號 !"); //不正確
				return true;
			}
			break;   
	}
}
</script>
