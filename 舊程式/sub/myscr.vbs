
<script Language=VBScript>
'==client
'pStr:��Ƥ��e
'pLen:��Ƴ̤j����,�Y�ǤJ0�h�Ǧ^��ƪ���
'pmsg:���W��,�YError�h�^�� ""
Function fDataLen(pStr,pLen,pmsg)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	fDataLen = 0
	tStr1 = ""
	if Len(pStr)=0 then tLen=0
	For ixI = 1 To Len(pStr)
		tStr1 = Mid(pStr, ixI, 1)
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then
			tLen = tLen + 2
		Else
			tLen = tLen + 1
		End If
	Next
	if pLen = 0 or tLen <= pLen then
		fDataLen = tLen
	else
		msgbox pmsg & "���׹L���A���ˬd!!!"
		fDataLen = ""
	end if
End Function
'pObj:�ˬd���פ�����
'pmsg:���W��,�YError�h�^�� ""
Function fChkDataLen(pObj,pmsg)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	Dim tc,te
	fChkDataLen = 0
	tStr1 = ""
	tLen = 0
	
	pObj.value = replace(pObj.value,"&","��")
	pObj.value = replace(pObj.value,"'","��")
	
	For ixI = 1 To Len(pObj.value)
		tStr1 = Mid(pObj.value, ixI, 1)
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then
			tLen = tLen + 2
		Else
			tLen = tLen + 1
		End If
	Next
	if pObj.maxlength = 0 or tLen <= pObj.maxlength then
		fChkDataLen = tLen
	else
		tc =  pObj.maxlength / 2
		te =  pObj.maxlength
		msgbox pmsg & " ���׹L���A���ˬd! " & chr(10) & chr(10) & "(����=����r�̦h: " & tc & "�Ӧr / �^��r�̦h: " & te & "�Ӧr)"
		pObj.focus
		fChkDataLen = ""
	end if
End Function
'check field null:�ˬd��@��줣�i���ť�
function chkNull(pFieldName,pobject)
	if trim(pobject.value)="" then
		msgbox pFieldName+"������J!!!"
		pobject.focus()
		chkNull = true
		exit function
	end if
	chkNull = false
End Function
'check field integer:�ˬd��@��줣�i���p��
function chkInt(pFieldName,pobject)
	if pobject.value > 0 then
	   tvalue=pobject.value	
	   tint=int(pobject.value)
           tvalue=tvalue / tint
	   if tvalue <> 1 then
	      msgbox pFieldName+"��������ơA�Э��s��J!!!"
  	      chkInt = true
	      exit function
	   end if
	end if
	chkInt = false
End Function

'******************************************

Function SetRadioValue(gObject,gd)
	for each x in gObject
		if x.value = gd then
			x.checked = true
		end if
	next
End Function

Function SetRadioValueNull(gObject)
	for each x in gObject
		x.checked = false
	next
End Function

Function GetRadioValue(gObject)
	rs = ""
	for each x in gObject
		if x.checked = true then
			rs = rs &  x.value & ";"
		end if
	next
        if rs <> "" then rs=left(rs,len(rs)-1)   
	GetRadioValue = rs
End Function

Function SetRadioDisabled(gObject)
	for each x in gObject
		x.disabled = true
	next
End Function


Function SetRsCode_Default (gs,gd)
	for each x in gs		
		if len(x.value) > 0 then
			gn = instr(x.value,"_")
			gvalue = mid(x.value,1,gn - 1)
			if gvalue = gd then
				x.selected = true
				exit function
			end if
		end if
	next
End function
</script>


<script langauge="vbscript">
'==date
function chkSEdate(psdate,pedate,pmsg)
	chkSEdate = true
	if psdate.value="" or pedate.value="" then
		alert(pmsg&"�_��������J!!!")
		exit function
	end if
	if psdate.value<>empty and pedate.value<>empty then
		if cdate(psdate.value)>cdate(pedate.value) then
			alert(pmsg&"�_�l�餣�i�j�󨴤��!!!")
			exit function
		end if
	end if
	chkSEdate = false
end function
'�ˬd����榡
function chkdateformat(pobject)
	chkdateformat = false
	if trim(pobject.value)=empty then exit function
	if isdate(pobject.value)=false then
		msgbox "����榡���~�A�Э��s��J!!! ����榡:YYYY/MM/DD"
		chkdateformat = true
		pobject.focus()
	else
		pobject.value = cdate(pobject.value)		
	end if
end function
function chkdateformat1(pobject,pmsg)
	chkdateformat1 = false
	if trim(pobject.value)=empty then exit function
	if isdate(pobject.value)=false then
		msgbox pmsg & "����榡���~�A�Э��s��J!!! ����榡:YYYY/MM/DD"
		pobject.focus()
		chkdateformat1 = true
	else
		pobject.value = cdate(pobject.value)		
	end if
end function
'����ɶ��榡
function getformatdatetime(pvalue,pkind)
	if trim(pvalue)=empty then exit function
'msgbox pkind
	tdate = year(pvalue) & "/" & string(2-len(month(pvalue)),"0") & month(pvalue) & "/" & string(2-len(day(pvalue)),"0") & day(pvalue)
	ttime = string(2-len(hour(pvalue)),"0") & hour(pvalue) & ":" & string(2-len(minute(pvalue)),"0") & minute(pvalue) & ":" & string(2-len(second(pvalue)),"0") & second(pvalue)
	select case pkind
		case "date"
			getformatdatetime = tdate
		case "time"
			getformatdatetime = ttime
		case "datetime"
			getformatdatetime = tdate & " " & ttime
	end select
'msgbox getformatdatetime
end function
'********************************

Function ChkDate(gObject)
	if trim(gObject.value)=empty then
		ChkDate = true 
		exit function
	end if
	if isdate(gObject.value)=false then
		msgbox "����榡���~�A�Э��s��J!!! ����榡:YYYY/MM/DD"
		gObject.focus()
		ChkDate = false
	else		
		ChkDate = true
	end if
End Function

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


<script Language=vbScript>
'=num
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
</script>