<script Language="javascript">
/*	chkSEDate: �_�l�餣�i�j�󨴤��
	chkEDate:	�ˬd�褸�~��J���T�_
*/
function chkSEDate(pSdate,pEdate) //�_�l�餣�i�j�󨴤��
{
	if (pSdate=="" && pEdate=="")
	{	return true;}
	if (ChkEDate(pSdate)[1] > ChkEDate(pEdate)[1])
	{	alert("�_�l�餣�i�j�󨴤��");
		return true;}
	return false;
}
function ChkEDate(pDateStr)  //�s�f�褸�~
{
//�ǤJ�Ѽ�:YYYYMMDD(ex:19990101) or YYYY/MM/DD(ex:1999/01/01 or 1999/1/1)
//�Ǧ^�Ѽ�:Error:Return Array(true,""), Success:Return Array(false,�зǮ榡�褸�~)
	if (pDateStr == "")
	{   return Array(false,pDateStr);}
	tPos = pDateStr.indexOf("/");
	if (tPos == -1)   //��J19990101 (��J����ƭY���t/,����J8��)
	{
		if (pDateStr.length != 8) 
		{
			alert(pDateStr+"����榡���~ (ex:YYYYMMDD) !");
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
	else         //��J1999/01/01 or ��J1999/1/1
	{
		tYYYY = parseInt(pDateStr.substring(0,tPos),10) //�~
		if (tYYYY.toString().length != 4)
		{
			alert("����榡���~ (ex:YYYY/MM/DD) !");
			return Array(true,"");
		}
		tDateStr = pDateStr.substring(tPos+1);
		tPos   = tDateStr.indexOf("/");
		tMM    = tDateStr.substring(0,tPos);      //��
		if (tMM.length == 1) {tMM = "0" + tMM;}
		tMM1   = parseInt(tMM,10) - 1;
		tDD    = tDateStr.substring(tPos+1);      //��
		if (tDD.length == 1) {tDD = "0" + tDD;}
	}
	if (tYYYY>=0 && tYYYY<=9999) {tTYYYY=tYYYY;}
	if (tMM>=1 && tMM<=12) {tTMM=tMM;}
	if (tDD>=1 && tDD<=31) {tTDD=tDD;}
//	var tDate = new Date(tYYYY,tMM1,tDD).toLocaleString();    //�N����ରMM/DD/YYYY
	//�ݰt�Xserver�ɶ��]�w
//	tTYYYY = tDate.substring(6,10);
//	tTMM = tDate.substring(0,2);
//	tTDD = tDate.substring(3,5);
	if (tYYYY != tTYYYY || tMM != tTMM)
	{
		alert(pDateStr+"��J�L�Ī��褸��� !");
		return Array(true,"");
	}
	else
	{	tResult=tTYYYY+"/"+tMM+"/"+tDD;}
	return Array(false,tResult);
}
</script>