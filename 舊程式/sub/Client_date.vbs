<script Language="vbscript">
function chkSEdate(psdate,pedate,pmsg)
	chkSEdate = true
	if psdate="" or pedate="" then
		alert(pmsg&"起迄必須輸入!!!")
		exit function
	end if
	if psdate<>empty and pedate<>empty then
		if cdate(psdate)>cdate(pedate) then
			alert(pmsg&"起始日不可大於迄止日!!!")
			exit function
		end if
	end if
	chkSEdate = false
end function
function chkSEdate1(psdate,pedate,pmsg)
	chkSEdate1 = true
	if psdate="" or pedate="" then
		chkSEdate1 = false
		exit function
	end if
	if psdate<>empty and pedate<>empty then
		if cdate(psdate)>cdate(pedate) then
			alert(pmsg&"起始日不可大於迄止日!!!")
			exit function
		end if
	end if
	chkSEdate1 = false
end function
'檢查日期格式
function chkdateformat(pobject)
	chkdateformat = false
	if trim(pobject.value)=empty then exit function
	if isdate(pobject.value)=false then
		msgbox "日期格式錯誤，請重新輸入!!! 日期格式:YYYY/MM/DD"
		pobject.focus()
		chkdateformat = true
	else
		pobject.value = cdate(pobject.value)
	end if
end function
'檢查日期格式 YYYYMM
function chkdateformat1(pvalue,pmsg)
	chkdateformat1 = false
	if trim(pvalue)=empty then exit function
	pvalue = left(pvalue,4) & "/" & mid(pvalue,5,2) & "/01"
	if isdate(pvalue)=false then
		msgbox pmsg&"日期格式錯誤，請重新輸入!!! 日期格式:YYYYMM"
		chkdateformat1 = true
	end if
end function

function ChkDateSE(psdate,pedate,pmsg)
	ChkDateSE = true
	if psdate="" or pedate="" then
		ChkDateSE = true
		exit function
	end if
	if psdate<>empty and pedate<>empty then
		if cdate(psdate)>cdate(pedate) then
			alert(pmsg&"起始日不可大於迄止日!!!")
			ChkDateSE = false
			exit function
		end if
	end if
	ChkDateSE = true
end function
</script>
