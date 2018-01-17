<script Language="javascript">
/*	getselect:	將selected停於該筆資料
	chkname:	輸入名稱檢查是否存在於combobox裡,並將combobox停於該筆資料
*/
//show combobox值
function getselect(pObject,pValue) {
	if (pValue != "") {
		for (q=0;q<pObject.length-1;q++) {
			if (pObject.options(q).value==pValue)
			{	pObject.options(q).selected=true;}
		}
	}
}
//check姓名輸入在select裡否
function chkname(pObject,pValue) {
	for (q=1;q<pObject.length-1;q++) {
		var tPos = pObject.options(i).text.indexOf("_");
		if (pObject.options(q).text.substring(tPos+1)==pValue) {
			getselect(pObject,pObject.options(q).value);
			return false;
		}
	}
	alert("無"+pValue+"姓名，請檢查是否輸入錯誤或資料庫無建此資料!!!");
	return true;
}
</script>
