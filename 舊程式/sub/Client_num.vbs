<script Language=vbScript>
function chkNum(pvalue,pmsg)
	chkNum = false
	if pvalue<>empty then 
		if IsNumeric(pvalue) = false then
			msgbox pmsg & "�������ƭ�!!!"
			chkNum = true
			exit function
		end if
	end if
end function
function chkNum1(pobj,pmsg)
	chkNum1 = false
	if pobj.value<>empty then 
		if IsNumeric(pobj.value) = false then
			msgbox pmsg & "�������ƭ�!!!"
			pobj.focus()
			chkNum1 = true
			exit function
		end if
	end if
end function
function chkSENum(ps,pe,pmsg)
	chkSENum = false
	if ps.value<>empty and pe.value<>empty then 
		if cdbl(ps.value)>cdbl(pe.value) then
			msgbox pmsg & " �_�l�s�����i�j�󨴤�s���������ƭ�!!!"
			ps.focus()
			chkSENum = true
			exit function
		end if
	end if
end function
</script>
