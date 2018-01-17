<script language="vbscript">
function getoption(pobject,pvalue)'combobox停在正確資料上
	for w=0 to pobject.length-1
		if pobject(w).value=trim(pvalue) then
			pobject(w).selected=true
		end if
	next
end function

Function getradiovalue(pobject)
        getradiovalue = ""
        For i=0 To pobject.length-1
            If pobject(i).checked = True Then
                getradiovalue = pobject(i).value
                Exit For
            End If
        Next
End Function
</script>
