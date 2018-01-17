<script Language=vbScript>
function chkNum(pvalue,pmsg)
	chkNum = false
	if pvalue<>empty then 
		if IsNumeric(pvalue) = false then
			msgbox pmsg & "必須為數值!!!"
			chkNum = true
			exit function
		end if
	end if
end function
function chkNum1(pobj,pmsg)
	chkNum1 = false
	if pobj.value<>empty then 
		if IsNumeric(pobj.value) = false then
			msgbox pmsg & "必須為數值!!!"
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
			msgbox pmsg & " 起始編號不可大於迄止編號必須為數值!!!"
			ps.focus()
			chkSENum = true
			exit function
		end if
	end if
end function
</script>
