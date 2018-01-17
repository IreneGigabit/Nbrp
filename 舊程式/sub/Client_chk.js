<script Language=javaScript>
/*	chkNull:	檢查單一欄位不可為空白
	archkNull:	檢查多個欄位不可為空白
*/
//check field null
function chkNull(pFieldName,pobject)
{
	if (pobject.value=="") {
		alert(pFieldName+"必須輸入!!!");
		pobject.focus();
		return true;
	}
	return false;
}
//check field null
function archknull(pnm,pva)
{
	nm = pnm.split("|");
	va = pva.split("|");
	for (z=0;z<nm.length;z++) {
		if (va[z]=="") {
			alert(nm[z]+"必須輸入!!!");
			return true;
		}
	}
	return false;
}
</script>
