<script Language="vbscript">
function chkSEdate(psdate,pedate,pmsg)
	chkSEdate = true
	if psdate="" or pedate="" then
		alert(pmsg&"�_��������J!!!")
		exit function
	end if
	if psdate<>empty and pedate<>empty then
		if cdate(psdate)>cdate(pedate) then
			alert(pmsg&"�_�l�餣�i�j�󨴤��!!!")
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
			alert(pmsg&"�_�l�餣�i�j�󨴤��!!!")
			exit function
		end if
	end if
	chkSEdate1 = false
end function
'�ˬd����榡
function chkdateformat(pobject)
	chkdateformat = false
	if trim(pobject.value)=empty then exit function
	if isdate(pobject.value)=false then
		msgbox "����榡���~�A�Э��s��J!!! ����榡:YYYY/MM/DD"
		pobject.focus()
		chkdateformat = true
	else
		pobject.value = cdate(pobject.value)
	end if
end function
'�ˬd����榡 YYYYMM
function chkdateformat1(pvalue,pmsg)
	chkdateformat1 = false
	if trim(pvalue)=empty then exit function
	pvalue = left(pvalue,4) & "/" & mid(pvalue,5,2) & "/01"
	if isdate(pvalue)=false then
		msgbox pmsg&"����榡���~�A�Э��s��J!!! ����榡:YYYYMM"
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
			alert(pmsg&"�_�l�餣�i�j�󨴤��!!!")
			ChkDateSE = false
			exit function
		end if
	end if
	ChkDateSE = true
end function
</script>
