<script Language="VBScript">
'---�̰�O�a�i�M���קO
function getcoun_pbr_type(pagent_no,pcountry)
	'update agent set pbr_type='01' where agcountry='JA'-- �饻��
	'update agent set pbr_type='02' where agcountry='AM' and agent_no<>'D279'--����� �D HP
	'update agent set pbr_type='03' where pbr_type is null -- HP�פΫD�����
	if pcountry="JA" then
		getcoun_pbr_type = "01" 
	elseif pcountry="AM" and pagent_no<>"D279" then	
		getcoun_pbr_type = "02"
	else
		getcoun_pbr_type = "03"
	end if
end function

function IMG_Click(pType)
	if reg.rs_kind(0).checked=true then
		rs_kind = "Y"
		reg.agent_no.value = ucase(reg.agent_no.value)
		reg.agent_no1.value = ucase(reg.agent_no1.value)
		tvalue = reg.agent_no.value
	else
		rs_kind = "N"
		tvalue = reg.agt_sqlno.value
	end if
	if tvalue<>empty then
		select case pType
			case "1" '�����ި�
				window.open "Agent34Edit.asp?submitTask=Q&qtype=N&closewin=Y&prgid=" &reg.prgid.value& "&rs_kind="& rs_kind &"&agt_sqlno=" &reg.agt_sqlno.value& "&Agent_no=" &reg.Agent_no.value& "&Agent_no1=" &reg.Agent_no1.value,"myWindowOne", "width=800 height=600 top=40 left=80 toolbar=no, menubar=no, location=no, directories=no resizeable=no status=no scrollbars=yes"
			case "2" '���o�i��
				window.open "Agent33list.asp?submitTask=Q&closewin=Y&prgid=" &reg.prgid.value& "&rs_kind="& rs_kind& "&qagt_sqlno=" &reg.agt_sqlno.value& "&qagent_no=" &reg.agent_no.value& "&qagent_no1=" &reg.agent_no1.value,"myWindowOne", "width=800 height=600 top=40 left=80 toolbar=no, menubar=no, location=no, directories=no resizeable=no status=no scrollbars=yes"
		end select
	else 
		msgbox "�Х���J�N�z�H�s����A����d�ߥ\��!!"
		if rs_kind = "Y" then
			reg.agent_no.focus
		else
			reg.agt_sqlno.focus
		end if
		exit function
	end if
end function
function IMG_Click1(pagt_sqlno,pagent_no,pagent_no1,pType,prs_kind)
	pagent_no = ucase(pagent_no)
	pagent_no1 = ucase(pagent_no1)
	select case pType
		case "1" '�����ި�
			window.open "../agent3m/Agent34Edit.asp?submitTask=Q&qtype=N&closewin=Y&prgid=agent34&rs_kind="& prs_kind &"&agt_sqlno=" &pagt_sqlno& "&Agent_no=" &pagent_no& "&Agent_no1=" &pagent_no1,"myWindowOne", "width=800 height=600 top=40 left=80 toolbar=no, menubar=no, location=no, directories=no resizeable=no status=no scrollbars=yes"
		case "2" '���o�i��
			window.open "../agent3m/Agent33list.asp?submitTask=Q&closewin=Y&prgid=agent33&rs_kind="& prs_kind &"&qagt_sqlno=" &pagt_sqlno& "&qagent_no=" &pagent_no& "&qagent_no1=" &pagent_no1,"myWindowOne", "width=800 height=600 top=40 left=80 toolbar=no, menubar=no, location=no, directories=no resizeable=no status=no scrollbars=yes"
	end select
end function

'pStr:��Ƥ��e
'pLen:��Ƴ̤j����,�Y�ǤJ0�h�Ǧ^��ƪ���
'pmsg:���W��,�YError�h�^�� ""
Function fDataLen(pStr,pLen,pmsg,pa)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	fDataLen = 0
	tStr1 = ""
'	if Len(pStr)=0 then tLen=0
	pStr = replace(pStr,"&","��")
	pStr = replace(pStr,"'","��")
	pStr = replace(pStr,"""","��")
	pa.value  = pStr
	
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
		pa.focus
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
	pObj.value = replace(pObj.value,"""","��")
	
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

</script>
