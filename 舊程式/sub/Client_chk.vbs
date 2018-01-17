<script Language=VBScript>
function fseq_chk(pobject)
	if pobject.value <> "" then
       b=trim(pobject.value)
       if right(b,1) = "," then 
          alert "�̫᪺���ҽs�����i���r��!"
          pobject.focus()
          exit function
       end if   
                    
       b1=split(b,",")
       for i=0 to ubound(b1)
           if not isnumeric(b1(i)) then 
              alert "�y"&b1(i)&"�z�ݬ��ƭ�!"
              pobject.focus()
              exit function
           end if   
       next 
    end if
end function
'pStr:��Ƥ��e,pLen:��Ƴ̤j����,�Y�ǤJ0�h�Ǧ^��ƪ���,pmsg:���W��,�YError�h�^�� ""
Function fDataLen(pStr,pLen,pmsg)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	fDataLen = 0
	tLen=0
	tStr1 = ""
	if Len(pStr)=0 then
		fDataLen = "0"
		exit function
	end if
	For ixI = 1 To Len(pStr)
		tStr1 = Mid(pStr, ixI, 1)
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then
			tLen = tLen + 2
		Else
			tLen = tLen + 1
		End If
	Next
'msgbox tLen
	if pLen=0 or tLen<=pLen then
		fDataLen = tLen
	else
		msgbox pmsg & "���׹L���A���ˬd!!! (�̦h�u�i��J" & pLen/2 & "�Ӥ����" & pLen & "�ӭ^�Ʀr)"
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
function chkNull1(pFieldName,pvalue)
	if trim(pvalue)="" then
		msgbox pFieldName+"������J!!!"
		chkNull1 = true
		exit function
	end if
	chkNull1 = false
End Function
'�Y�ť���ܴ���
function chkNull2(pFieldName,pvalue)
	chkNull2 = false
	if trim(pvalue)="" then
		'ans = msgbox("�u"& pFieldName &"�v��ƪť�!!!"&chr(10)&chr(10)&"�Y�~�����@�~�Ы��u�O�v�A�Y�n�ק��ƽЫ��u�_�v!!!",vbYesNo,"����.....")
		msgbox "�`�N:�u"& pFieldName &"�v��ƪť�!!!"
		'if ans = 6 then
		'else
		'	chkNull2 = true
		'	exit function
		'end if
	end if
End Function
'check field integer:�ˬd��@��줣�i���p��
function chkInt(pFieldName,pobject)
	if IsNumeric(pobject.value)=true then
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
	else
		msgbox pFieldName+"��������ơA�Э��s��J!!!"
  		chkInt = true
		exit function
	end if
	chkInt = false
End Function
'�ˬd�ި����,�O�_�ݭn�{���鴣�e�ܶg��
function check_week_ctrldate(pctrl_type,pdate)
	Dim gCtrl_type,tCtrl_Type , tday , tdate,tbaseday
	
	'msgbox "pdate=" & pdate
	'msgbox "pctrl_type=" & pctrl_type
	
	if (pdate="") or isnull(pdate) then
		check_week_ctrldate = pdate
		exit function
	end if
	
	'�ثe�u���ۺ޴���,�ӿ����,�ӿ�o������ݭn�{���鴣�e�ܶg��
	gCtrl_Type = "B,D,F"
	
	tctrl_type = left(pctrl_type,1)

	if instr(gCtrl_type,tCtrl_Type)=0 then
		check_week_ctrldate = pdate
		exit function
	end if

	'tday: 1:�P���@ 2:�P���G 3:�P���T 4:�P���| 5:�P���� 6:�P���� 7:�P����
	
	tday = datepart("W",pdate,2)
	
	'�g��,�g��~�ݳB�z
	if tday=6 or tday=7 then
		tbaseday=5-tday
		check_week_ctrldate=dateadd("d",tbaseday,pdate)
	else
		check_week_ctrldate=pdate
	end if
	
end function
Function fRound(pnum,pdec)
	'pnum�Ʀr
	'pdec��ܨ�p�ƴX��
	dim tdec
	dim tnum
	
	pdec=cint(pdec)
	pnum = cdbl(pnum)
	'formatnumber(�Ʀr,�p�Ʀ��,���I�e�O�_��ܫe�ɹs,�t�ƭȬO�_�a���A��,�O�_�H�Ʀ�s�ղŸ��Ӥ��j)
	fround=formatnumber(pnum,pdec,,,0)
End Function

'�N URL ���S��r��������
Function fURLEncode(pdata)
    dim tdata
    dim fsql    
    tdata = pdata
    
    fsql = "select * from vURLEncode "
    fsql = fsql & " where code_type = 'purlencode' "
    fsql = fsql & " order by sortfld "
    url = "../xml/XmlGetSqlDataMulti.asp?searchsql=" & fsql
    Set xmldoc = CreateObject("Microsoft.XMLDOM")
    xmldoc.async = false
    If xmldoc.load (url) then
	    Set root = xmldoc.documentElement
	    For Each xi In root.childNodes
            tdata = replace(tdata,trim(xi.selectsinglenode("cust_code").text),trim(xi.selectsinglenode("code_name").text))
	    Next
	    Set root = Nothing
    else
        msgbox "���o URLEncode���ѡI"
        exit function
    End If
    Set xmldoc = Nothing
    fURLEncode = tdata
end function

'�w��textbox��ܪ����e�A�n�����S��r��
function txtCnv(pStr)
	Dim tStr
	
	if isnull(pStr) or pStr = "" then
		txtCnv = ""
	else
		tStr = pStr
		tStr = replace(tStr,"'","��")
		tStr = replace(tStr,"""","��")
		tStr = replace(tStr,"&","��")
		txtCnv = tStr
	end if
end function

function OpenField(preg)
	dim x
	for each x in preg
		select case x.type		
			case "select-one"
				x.disabled = false
			case "textarea"
				x.disabled = false
			case "radio"
				x.disabled = false
			case "checkbox"
				x.disabled = false
			case "text"
				x.disabled = false
		end select
	next
end function

Function chkRadio(pobject, pmsg)
    Dim sret
    sret = False
    
    For i=0 To pobject.length-1
        If pobject(i).checked=True Then
            sret = True
            Exit For
        End If
    Next
    
    If sret=False Then
        MsgBox "�п��"& pmsg & "�I"
    End If
    
    chkRadio = sret
End Function

Function chkTest_onclick()
	If reg.chkTest.checked=True Then	
		document.getElementById("ActFrame").style.display = ""
	Else
		document.getElementById("ActFrame").style.display = "none"
	End If
End Function

Function IIf(bBool, trueStr, falseStr)
	If bBool Then
		IIf = trueStr
	Else 
		IIf = falseStr
	End If
End Function

</script>

