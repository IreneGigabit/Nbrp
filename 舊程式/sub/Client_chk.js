<script Language=javaScript>
/*	chkNull:	�ˬd��@��줣�i���ť�
	archkNull:	�ˬd�h����줣�i���ť�
*/
//check field null
function chkNull(pFieldName,pobject)
{
	if (pobject.value=="") {
		alert(pFieldName+"������J!!!");
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
			alert(nm[z]+"������J!!!");
			return true;
		}
	}
	return false;
}
</script>
