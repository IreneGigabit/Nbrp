<script Language=javaScript>
function archkNaN(pnm,pva) //����J�������ƭ�
{
	nm = pnm.split("|");
	va = pva.split("|");
	for (z=0;z<nm.length;z++) {
		if (va[z]!="")
		{	if (isNaN(va[z]))
			{	alert(nm[z]+"�������ƭ�!!!");
				return true;
			}
		}
	}
	return false;
}
function ChkNumValue(pNumStr,pChkZero,pChkInt) //�s�f�ƭ���쪺��J��
{   
/*	�ǤJ�Ѽ�:	1.pNumStr:�ƭ�
				2.pChkZero=0:��Check���ȥ��t��,1:Check���ȥ����j��0,2:Check���ȥ���>=0
				3.pChkInt=0:��Check���Ȭ����,1:Check���Ȭ����
	�Ǧ^�Ѽ�:	Error=true, Success=false	
*/
	if (pNumStr == "") {return false;}
	if (isNaN(pNumStr) == true) {
		alert("���ȶ����ƭ� !");
		return true;
	}
	switch(pChkZero)
	{
		case 0:
			break;
		case 1:
			if (pNumStr <= 0) {
				alert("�ƭȥ����j�� 0 !");
				return true;
			}
			break;
		case 2:
			if (pNumStr < 0) {
				alert("�ƭȥ����j�󵥩� 0 !");
				return true;
			}
			break;
	}
    if (pChkInt == 1) {
		if (parseInt(pNumStr) != pNumStr) {
			alert("���ȶ������ !");
			return true;
		}
    }
    return false;
}
function ChkNumFormat(pNumStr,pDefineInt,pPoint,pType)
{
/*�\�໡��:	pType=0:�|�ˤ��J��p�Ʋ�pPoint��,1:�L����˥h��p�Ʋ�pPoint��,
					2:�L����i���p�Ʋ�pPoint��
	�ǤJ�Ѽ�:	1.pNumStr:�ƭ�,
				2.pDefineInt:Table Layout�ҭq������Ƴ�������(1,2....)
				3.pPoint:Table Layout�ҭq�����p�Ƴ�������, ��Y�p���p�Ʋ�N��(0,1,2....),
				4.pType:0:�|�ˤ��J,1:�L����˥h,2:�L����i��
	�Ǧ^�Ѽ�:Error:Return Array(true,""),No Error: Return Array(false,�ഫ���T���Ʀr)
*/
	var tPointNum,tTempNum,tResult;
	var tPos,ixI;
	var swFmtErr;

	if (pNumStr == "") {return Array(false,pNumStr);}
	swFmtErr=false;
	if (isNaN(pNumStr) == true) //�T�w��J�ȬO�Ʀr
	{   alert("��J�Ȥ��O�Ʀr !");
		return Array(true,"");
	}
	//�T�w��J�Ȯ榡���T(ex:��J�Ȫ���Ʀ�ƲŦXTable Define)
	tPos = pNumStr.indexOf(".");
	if (tPos > 0)
	{	tIntNumber=pNumStr.substring(0,tPos);
		tPointNumber = pNumStr.substring(tPos+1,pNumStr.length);
		if (tIntNumber.length > pDefineInt)	{swFmtErr=true;}
	}
	else
	{   if (pNumStr.length > pDefineInt) {swFmtErr=true;}}
	if (swFmtErr==true)
	{   alert("��J�榡���~ !");
		return Array(true,"");
	}

	tPointNum=1;

	//��p���In��N��n��0
	for(var ixI=0;ixI<pPoint;ixI++) {tPointNum = tPointNum * 10;}
	tTempNum = (pNumStr * tPointNum);
	switch(pType)
	{
		case 0: //�|�ˤ��J
			tTempNum = Math.round(tTempNum);
			break;
		case 1: //�L����˥h
			tTempNum = parseInt(tTempNum);
			break;
		case 2: //�L����i��
			tTempNum = parseInt(tTempNum) + 1;
			break;
	}
	if (tPos > 0) //�Y�O���p�ƧP�_�p�Ʀ�ƬO�_�ŦX�P�_��
	{
		if (tPointNumber.length < pPoint)
		{   tResult = pNumStr;}
		else
		{   tResult = (tTempNum/tPointNum);}
	}
	else
	{   tResult=pNumStr;}  //�Y�O��ƶǾ�Ʀ^�h
	return Array(false,tResult);
}
</script>
