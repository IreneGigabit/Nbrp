<script Language=VBScript>
function fseq_chk(pobject)
	if pobject.value <> "" then
       b=trim(pobject.value)
       if right(b,1) = "," then 
          alert "最後的本所編號不可有逗號!"
          pobject.focus()
          exit function
       end if   
                    
       b1=split(b,",")
       for i=0 to ubound(b1)
           if not isnumeric(b1(i)) then 
              alert "『"&b1(i)&"』需為數值!"
              pobject.focus()
              exit function
           end if   
       next 
    end if
end function
'pStr:資料內容,pLen:資料最大長度,若傳入0則傳回資料長度,pmsg:欄位名稱,若Error則回傳 ""
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
		msgbox pmsg & "長度過長，請檢查!!! (最多只可輸入" & pLen/2 & "個中文或" & pLen & "個英數字)"
		fDataLen = ""
	end if
End Function
'pObj:檢查長度之物件
'pmsg:欄位名稱,若Error則回傳 ""
Function fChkDataLen(pObj,pmsg)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	Dim tc,te
	fChkDataLen = 0
	tStr1 = ""
	tLen = 0
	
	pObj.value = replace(pObj.value,"&","＆")
	pObj.value = replace(pObj.value,"'","’")
	
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
		msgbox pmsg & " 長度過長，請檢查! " & chr(10) & chr(10) & "(提示=中文字最多: " & tc & "個字 / 英文字最多: " & te & "個字)"
		pObj.focus
		fChkDataLen = ""
	end if
End Function

'check field null:檢查單一欄位不可為空白
function chkNull(pFieldName,pobject)
	if trim(pobject.value)="" then
		msgbox pFieldName+"必須輸入!!!"
		pobject.focus()
		chkNull = true
		exit function
	end if
	chkNull = false
End Function
function chkNull1(pFieldName,pvalue)
	if trim(pvalue)="" then
		msgbox pFieldName+"必須輸入!!!"
		chkNull1 = true
		exit function
	end if
	chkNull1 = false
End Function
'若空白顯示提示
function chkNull2(pFieldName,pvalue)
	chkNull2 = false
	if trim(pvalue)="" then
		'ans = msgbox("「"& pFieldName &"」資料空白!!!"&chr(10)&chr(10)&"若繼續執行作業請按「是」，若要修改資料請按「否」!!!",vbYesNo,"提示.....")
		msgbox "注意:「"& pFieldName &"」資料空白!!!"
		'if ans = 6 then
		'else
		'	chkNull2 = true
		'	exit function
		'end if
	end if
End Function
'check field integer:檢查單一欄位不可為小數
function chkInt(pFieldName,pobject)
	if IsNumeric(pobject.value)=true then
		if pobject.value > 0 then
		   tvalue=pobject.value	
		   tint=int(pobject.value)
		       tvalue=tvalue / tint
		   if tvalue <> 1 then
		      msgbox pFieldName+"必須為整數，請重新輸入!!!"
  		      chkInt = true
		      exit function
		   end if
		end if
	else
		msgbox pFieldName+"必須為整數，請重新輸入!!!"
  		chkInt = true
		exit function
	end if
	chkInt = false
End Function
'檢查管制期限,是否需要逢假日提前至週五
function check_week_ctrldate(pctrl_type,pdate)
	Dim gCtrl_type,tCtrl_Type , tday , tdate,tbaseday
	
	'msgbox "pdate=" & pdate
	'msgbox "pctrl_type=" & pctrl_type
	
	if (pdate="") or isnull(pdate) then
		check_week_ctrldate = pdate
		exit function
	end if
	
	'目前只有自管期限,承辦期限,承辦發文期限需要逢假日提前至週五
	gCtrl_Type = "B,D,F"
	
	tctrl_type = left(pctrl_type,1)

	if instr(gCtrl_type,tCtrl_Type)=0 then
		check_week_ctrldate = pdate
		exit function
	end if

	'tday: 1:星期一 2:星期二 3:星期三 4:星期四 5:星期五 6:星期六 7:星期日
	
	tday = datepart("W",pdate,2)
	
	'週六,週日才需處理
	if tday=6 or tday=7 then
		tbaseday=5-tday
		check_week_ctrldate=dateadd("d",tbaseday,pdate)
	else
		check_week_ctrldate=pdate
	end if
	
end function
Function fRound(pnum,pdec)
	'pnum數字
	'pdec顯示到小數幾位
	dim tdec
	dim tnum
	
	pdec=cint(pdec)
	pnum = cdbl(pnum)
	'formatnumber(數字,小數位數,數點前是否顯示前導零,負數值是否帶有括號,是否以數位群組符號來分隔)
	fround=formatnumber(pnum,pdec,,,0)
End Function

'將 URL 中特殊字元替換掉
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
        msgbox "取得 URLEncode失敗！"
        exit function
    End If
    Set xmldoc = Nothing
    fURLEncode = tdata
end function

'針對textbox顯示的內容，要替換特殊字元
function txtCnv(pStr)
	Dim tStr
	
	if isnull(pStr) or pStr = "" then
		txtCnv = ""
	else
		tStr = pStr
		tStr = replace(tStr,"'","’")
		tStr = replace(tStr,"""","”")
		tStr = replace(tStr,"&","＆")
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
        MsgBox "請選擇"& pmsg & "！"
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

