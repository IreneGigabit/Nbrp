<script Language="javascript">
/*	getselect:	�Nselected����ӵ����
	chkname:	��J�W���ˬd�O�_�s�b��combobox��,�ñNcombobox����ӵ����
*/
//show combobox��
function getselect(pObject,pValue) {
	if (pValue != "") {
		for (q=0;q<pObject.length-1;q++) {
			if (pObject.options(q).value==pValue)
			{	pObject.options(q).selected=true;}
		}
	}
}
//check�m�W��J�bselect�̧_
function chkname(pObject,pValue) {
	for (q=1;q<pObject.length-1;q++) {
		var tPos = pObject.options(i).text.indexOf("_");
		if (pObject.options(q).text.substring(tPos+1)==pValue) {
			getselect(pObject,pObject.options(q).value);
			return false;
		}
	}
	alert("�L"+pValue+"�m�W�A���ˬd�O�_��J���~�θ�Ʈw�L�ئ����!!!");
	return true;
}
</script>
