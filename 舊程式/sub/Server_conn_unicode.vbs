<%
' 把文字變成Unicode字串 (寫入 DB nvarchar ntext 欄位時要先將 anc 轉成 unicode)
Function ToUnicode(s)
    Dim ret,i,c,a,w
	    
  	If IsNull(s) or Trim(s)="" then
		ToUnicode = ""
		Exit Function
	End If 
		  
    Set regEx = New RegExp
    regEx.Pattern = "&#([0-9]+);"
    regEx.IgnoreCase = True
    regEx.Global = True
	    
    ret = s
	    
    'Response.Write chrw(28286) & "<br>"
    'Response.Write chrB(28286) & "<br>"
    'Response.Write asc(28286) & "<br>"
    'Response.Write "chrw(28286):"& chrw(28286) & "<br>"
	    
    'Response.Write "原資料有無&#:<br>"& ret & "<BR>"
    If regEx.Test(s) then
        Set Matches = regEx.Execute(s)
        for Each Match in Matches
            repStr = Mid(Match.value,3,Len(Match.value)-3)
            'Response.Write "&#"& repStr &";" & "---"& repStr &"<br>"
            ret = Replace(ret,"&#"& repStr &";",chrw(repStr))
            'Response.Write chrw(repStr) & "<br>"
        next
        'Response.Write "Replace後資料:<br>"& ret & "<BR><BR>"
    End If
    ToUnicode = ret
End Function

Function ToUnicode2(s)
    Dim ret,i,c,a,w
	    
  	If IsNull(s) then
		ToUnicode2 = ""
		Exit Function
	End If 
		  
    Set regEx = New RegExp
    regEx.Pattern = "&#([0-9]+);"
    regEx.IgnoreCase = True
    regEx.Global = True
	    
    ret = s
	    
    'Response.Write chrw(28286) & "<br>"
    'Response.Write chrB(28286) & "<br>"
    'Response.Write asc(28286) & "<br>"
    'Response.Write "chrw(28286):"& chrw(28286) & "<br>"
	    
    'Response.Write "原資料有無&#:<br>"& ret & "<BR>"
    If regEx.Test(s) then
        Set Matches = regEx.Execute(s)
        for Each Match in Matches
            repStr = Mid(Match.value,3,Len(Match.value)-3)
            'Response.Write "&#"& repStr &";" & "---"& repStr &"<br>"
            ret = Replace(ret,"&#"& repStr &";",chrw(repStr))
            'Response.Write chrw(repStr) & "<br>"
        next
        'Response.Write "Replace後資料:<br>"& ret & "<BR><BR>"
    End If
    ToUnicode2 = ret
End Function

Function ToXmlUnicode(s)
    Dim ret,i,c,a,w
	    
  	If IsNull(s) then
		ToXmlUnicode = ""
		Exit Function
	End If
	
	s=replace(s,"&","&amp;")
	
    Set regEx = New RegExp
    regEx.Pattern = "&amp;#([0-9]+);"
    regEx.IgnoreCase = True
    regEx.Global = True
	    
    ret = s
	    
    'Response.Write "原資料有無&#:<br>"& ret & "<BR>"
    If regEx.Test(s) then
		'Response.Write "有match:"& ret & "<BR>"
        Set Matches = regEx.Execute(s)
        for Each Match in Matches
            repStr = Mid(Match.value,7,Len(Match.value)-7)
            'Response.Write "&amp;#"& repStr &";" & "---"& repStr &"<br>"
            'ret = Replace(ret,"&#"& repStr &";",chrw(repStr))
            ret = Replace(ret,"&amp;#"& repStr &";","&#"& repStr &";")
            'Response.Write chrw(repStr) & "<br>"
        next
    End If
	
	ret=replace(ret,"＆","&amp;")
	ret=replace(ret,"<","&lt;")
	'ret=replace(ret,">","&gt;")
	'ret=replace(ret,"'","&apos;")
	'ret=replace(ret,"""","&quot;")
	
    ToXmlUnicode = ret
End Function

function chknull_unicode(p1)
	if trim(p1)<>empty then
		chknull_unicode = "N'" & ToUnicode(p1) & "'"
	else
		chknull_unicode = "null"
	end if
end function
function chkempty_unicode(p1)
	if trim(p1)<>empty then
		chkempty_unicode = "N'" & ToUnicode(p1) & "'"
	else
		chkempty_unicode = "''"
	end if
end function
function chkempty_unicode2(p1)
	if p1<>empty then
		chkempty_unicode2 = "N'" & ToUnicode2(p1) & "'" '沒有trim
	else
		chkempty_unicode2 = "''"
	end if
end function

'新增 Log 檔，適用於 log table 中有 ud_flag、ud_date、ud_scode、prgid 這些欄位者
'ptable：ex:step_imp 要新增至 step_imp_log 則傳入  step_imp
'pkey_filed：key 值欄位名稱，如有多個欄位請用；隔開
'pkey_value：與 pkey_field 相互配合，如有多個欄位請用；隔開
function insert_log_table_unicode(pconn,pud_flag,pprgid,ptable,pkey_field,pkey_value)
	dim tisql
	dim tfield_str
	dim ar_key_field
	dim ar_key_value
	dim wsql
	dim ti
	
	set tRS = Server.CreateObject("ADODB.Recordset")
	
	tfield_str = ""
	
	tisql = "select b.name from sysobjects a, syscolumns b "
	tisql = tisql & " where a.id = b.id  and a.name = '" & ptable & "' and a.xtype='U' "
	tisql = tisql & " order by b.colid "
	tRS.open tisql,pconn,1,1
	while not tRS.eof
		tfield_str = tfield_str & tRS("name") & ","
		tRS.MoveNext
	wend
	
	tRS.close
	tfield_str = left(tfield_str,len(tfield_str) - 1)
	
	ar_key_field = ""
	ar_key_value = ""
	wsql = ""
	
	if instr(1,pkey_field,";") <> 0 then
		ar_key_field = split(pkey_field,";")
		ar_key_value = split(pkey_value,";")
		for ti = 0 to ubound(ar_key_field)
			wsql = wsql & " and " & ar_key_field(ti) & " = '" & ar_key_value(ti) & "' "
		next
	else
		wsql = " and " & pkey_field & " = '" & pkey_value & "' "
	end if
	
	tisql = "insert into " & ptable & "_log(ud_flag,ud_date,ud_scode,prgid," & tfield_str & ")"
	tisql = tisql & "select N'" & pud_flag & "',getdate(),N'" &session("scode") & "',"
	tisql = tisql & "N'" &pprgid& "'," & tfield_str
	tisql = tisql & " from " & ptable
	tisql = tisql & " where 1 = 1 "
	tisql = tisql & wsql
	'response.write tisql & "<br>"
	'response.end
	
	pconn.execute tisql
	
	set tRS = nothing
end function

'入檔時，char給空白
function chkcharnull_unicode(pvalue)
	if trim(pvalue)<>empty then
		pvalue = replace(pvalue,"'","’")
		pvalue = replace(pvalue,"""","”")
		pvalue = replace(pvalue,"&","＆")
		chkcharnull_unicode = "N'"& trim(pvalue) &"'"
	else
		chkcharnull_unicode = "''"
	end if
End Function
'入檔時，char給空白，不換雙引號
function chkcharnull2_unicode(pvalue)
	if trim(pvalue)<>empty then
		pvalue = replace(pvalue,"'","’")
		pvalue = replace(pvalue,"&","＆")
		pvalue = replace(pvalue,"<","＜")
		pvalue = replace(pvalue,">","＞")
		chkcharnull2_unicode = "N'"& trim(pvalue) &"'"
	else
		chkcharnull2_unicode = "''"
	end if
End Function

' 顯示Unicode字元(避免出現?) (由資料庫抓取資料後要顯示至畫面前要先轉碼)
Function Unicode2Htm(s)
	Dim ret,i,c,a,w
	If IsNull(s) or Trim(s)="" then
		Unicode2Htm = ""
		Exit Function
	End If  
	 
	ret = ""
	for i=1 to Len(s)
		c = Mid(s,i,1)
		a = Asc(c)
		w = Ascw(c)
		If w<0 then
			w = 65536 + w
		End If
		If a=63 and w<>63 then
			ret = ret & "&#" & w & ";"   
		ElseIf w>127 and w<256 then
			ret = ret & "&#" & w & ";"
		Else
			ret = ret & c
		End If
	next
	Unicode2Htm = ret
End Function

'====================================================================================================
'UTF-8跟BIG5互轉
'====================================================================================================
'由資料庫抓取資料後要顯示至畫面前要先轉碼
Function UnicodeToBig5(str)
	Dim old,new_w,j
	old = str
	new_w = ""
	IF isnull(str) then
		UnicodeToBig5 = str
	Else
		For j = 1 To Len(str)
			if ascw(mid(old,j,1)) < 0 then
				new_w = new_w & "&#" & ascw(mid(old,j,1)) + 65536 & ";"
			ElseIf ascw(mid(old,j,1))>0 and ascw(mid(old,j,1))<127 then
				new_w = new_w & mid(old,j,1)
			Else
				new_w = new_w & "&#" & ascw(mid(old,j,1)) & ";"
			End if
		Next
		UnicodeToBig5 = new_w
	End IF	
End Function

'下列 function 有誤
'Function Big5ToUnicode(str) 
'	Dim x,y,z,temp_word,flag
'	flag = 0
'	x = InStr(flag + 1,str,"&#")
'	Do Until x = 0 or x < flag
'		x = InStr(flag + 1,str,"&#")
'		if x <> 0 then
'		y = Mid(str,x,8)
'		Select Case InStr(y,";")
'			Case 8
'				z = chrw(Mid(y,3,5))
'			Case 7
'				z = chrw(Mid(y,3,4))
'			Case 6
'				z = chrw(Mid(y,3,3))
'			Case 5
'				z = chrw(Mid(y,3,2))
'		End Select
'		if InStr(y,";") > 4 And Asc(z) <> 63 then
'			str = Replace(str,Left(y,InStr(y,";")),z)
'		End if
'			flag = x
'		End if
'	Loop
'	Big5ToUnicode = str
'End Function


Function getEmailDate(psend_date)
	getEmailDate=NumberToEngW(Weekday(psend_date)) & "," & NumberToEngaM(month(psend_date)) & " " & string(2-len(day(psend_date)),"0") & day(psend_date) & "," & year(psend_date) & " " & hour(psend_date) &":" & Minute(psend_date) & ":" & Second(psend_date) & " " &  NumberToEngaT(psend_date)
End Function




'月份數字轉大寫
Function NumberToEngaM(SendNumber)
   Select Case SendNumber
          Case "1"
               NumberToEngaM = "January"
          Case "2"
               NumberToEngaM = "February"
          Case "3"
               NumberToEngaM = "March"
          Case "4"
               NumberToEngaM = "April"
          Case "5"
               NumberToEngaM = "May"
          Case "6"
               NumberToEngaM = "June"
          Case "7"
               NumberToEngaM = "July"
          Case "8"
               NumberToEngaM = "August"
          Case "9"
               NumberToEngaM = "September"
          Case "10"
               NumberToEngaM = "October"
          Case "11"
               NumberToEngaM = "November"
          Case "12"
               NumberToEngaM = "December"
   End Select
End Function
'月份數字轉大寫(縮寫)
Function NumberToEngM(SendNumber)
   Select Case SendNumber
          Case "1"
               NumberToEngM = "Jan"
          Case "2"
               NumberToEngM = "Feb"
          Case "3"
               NumberToEngM = "Mar"
          Case "4"
               NumberToEngM = "Apr"
          Case "5"
               NumberToEngM = "May"
          Case "6"
               NumberToEngM = "Jun"
          Case "7"
               NumberToEngM = "Jul"
          Case "8"
               NumberToEngM = "Aug"
          Case "9"
               NumberToEngM = "Sep"
          Case "10"
               NumberToEngM = "Oct"
          Case "11"
               NumberToEngM = "Nov"
          Case "12"
               NumberToEngM = "Dec"
   End Select
End Function

'星期數字轉大寫
Function NumberToEngW(SendNumber)
   Select Case SendNumber
          Case "1"
               NumberToEngW = "Sunday"
          Case "2"
               NumberToEngW = "Monday"
          Case "3"
               NumberToEngW = "Tuesday"
          Case "4"
               NumberToEngW = "Wednesday"
          Case "5"
               NumberToEngW = "Thursday"
          Case "6"
               NumberToEngW = "Friday"
          Case "7"
               NumberToEngW = "Saturday"
   End Select
End Function
'星期數字轉大寫(縮寫)
Function NumberToEngaW(SendNumber)
   Select Case SendNumber
          Case "1"
               NumberToEngaW = "Sun"
          Case "2"
               NumberToEngaW = "Mon"
          Case "3"
               NumberToEngaW = "Tues"
          Case "4"
               NumberToEngaW = "Wed"
          Case "5"
               NumberToEngaW = "Thur"
          Case "6"
               NumberToEngaW = "Fri"
          Case "7"
               NumberToEngaW = "Sat"
   End Select
End Function
'星期數字轉大寫(縮寫)
Function NumberToEngaT(SendNumber)
   IF hour(SendNumber)>12 then
         NumberToEngaT = "PM"
   Else            
         NumberToEngaT = "AM"
   End IF
End Function

'for 轉換 xml 不接受之字元 ex.&
function fcdata(pdata)
    dim tdata
    
    tdata = pdata
    
    if isnull(tdata) = true then
        tdata = ""
    else
        tdata = replace(tdata,"&","&amp;")
    end if
    
    fcdata = tdata
end function

function HtmlCnv(pStr)
	Dim tStr
	
	if isnull(pStr) or pStr = "" then
		HtmlCnv = ""
	else
		tStr = pStr
		tStr = replace(tStr,"<","&lt;")
		tStr = replace(tStr,">","&gt;")
		tStr = replace(tStr,"\","&#34;")
		tStr = replace(tStr,"\r\n","&#13;&#10;")
		tStr = replace(tStr,"\n","&#13;&#10;")
		HtmlCnv = tStr
	end if
end function

'針對db入檔欄位替換特殊字元(特別針對附件檔名，因為附件檔名不可以直接換成全形的，因為這樣再查詢檢視時，會找不到原始檔案
function dbtxtCnv(pStr)
	Dim tStr
	
	if isnull(pStr) or pStr = "" then
		dbtxtCnv = ""
	else
		tStr = pStr
		tStr = replace(tStr,"'","''")
		dbtxtCnv = tStr
	end if
end function

%>
