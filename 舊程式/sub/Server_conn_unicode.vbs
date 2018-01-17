<%
' ���r�ܦ�Unicode�r�� (�g�J DB nvarchar ntext ���ɭn���N anc �ন unicode)
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
	    
    'Response.Write "���Ʀ��L&#:<br>"& ret & "<BR>"
    If regEx.Test(s) then
        Set Matches = regEx.Execute(s)
        for Each Match in Matches
            repStr = Mid(Match.value,3,Len(Match.value)-3)
            'Response.Write "&#"& repStr &";" & "---"& repStr &"<br>"
            ret = Replace(ret,"&#"& repStr &";",chrw(repStr))
            'Response.Write chrw(repStr) & "<br>"
        next
        'Response.Write "Replace����:<br>"& ret & "<BR><BR>"
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
	    
    'Response.Write "���Ʀ��L&#:<br>"& ret & "<BR>"
    If regEx.Test(s) then
        Set Matches = regEx.Execute(s)
        for Each Match in Matches
            repStr = Mid(Match.value,3,Len(Match.value)-3)
            'Response.Write "&#"& repStr &";" & "---"& repStr &"<br>"
            ret = Replace(ret,"&#"& repStr &";",chrw(repStr))
            'Response.Write chrw(repStr) & "<br>"
        next
        'Response.Write "Replace����:<br>"& ret & "<BR><BR>"
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
	    
    'Response.Write "���Ʀ��L&#:<br>"& ret & "<BR>"
    If regEx.Test(s) then
		'Response.Write "��match:"& ret & "<BR>"
        Set Matches = regEx.Execute(s)
        for Each Match in Matches
            repStr = Mid(Match.value,7,Len(Match.value)-7)
            'Response.Write "&amp;#"& repStr &";" & "---"& repStr &"<br>"
            'ret = Replace(ret,"&#"& repStr &";",chrw(repStr))
            ret = Replace(ret,"&amp;#"& repStr &";","&#"& repStr &";")
            'Response.Write chrw(repStr) & "<br>"
        next
    End If
	
	ret=replace(ret,"��","&amp;")
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
		chkempty_unicode2 = "N'" & ToUnicode2(p1) & "'" '�S��trim
	else
		chkempty_unicode2 = "''"
	end if
end function

'�s�W Log �ɡA�A�Ω� log table ���� ud_flag�Bud_date�Bud_scode�Bprgid �o������
'ptable�Gex:step_imp �n�s�W�� step_imp_log �h�ǤJ  step_imp
'pkey_filed�Gkey �����W�١A�p���h�����ХΡF�j�}
'pkey_value�G�P pkey_field �ۤ��t�X�A�p���h�����ХΡF�j�}
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

'�J�ɮɡAchar���ť�
function chkcharnull_unicode(pvalue)
	if trim(pvalue)<>empty then
		pvalue = replace(pvalue,"'","��")
		pvalue = replace(pvalue,"""","��")
		pvalue = replace(pvalue,"&","��")
		chkcharnull_unicode = "N'"& trim(pvalue) &"'"
	else
		chkcharnull_unicode = "''"
	end if
End Function
'�J�ɮɡAchar���ťաA�������޸�
function chkcharnull2_unicode(pvalue)
	if trim(pvalue)<>empty then
		pvalue = replace(pvalue,"'","��")
		pvalue = replace(pvalue,"&","��")
		pvalue = replace(pvalue,"<","��")
		pvalue = replace(pvalue,">","��")
		chkcharnull2_unicode = "N'"& trim(pvalue) &"'"
	else
		chkcharnull2_unicode = "''"
	end if
End Function

' ���Unicode�r��(�קK�X�{?) (�Ѹ�Ʈw�����ƫ�n��ܦܵe���e�n����X)
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
'UTF-8��BIG5����
'====================================================================================================
'�Ѹ�Ʈw�����ƫ�n��ܦܵe���e�n����X
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

'�U�C function ���~
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




'����Ʀr��j�g
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
'����Ʀr��j�g(�Y�g)
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

'�P���Ʀr��j�g
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
'�P���Ʀr��j�g(�Y�g)
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
'�P���Ʀr��j�g(�Y�g)
Function NumberToEngaT(SendNumber)
   IF hour(SendNumber)>12 then
         NumberToEngaT = "PM"
   Else            
         NumberToEngaT = "AM"
   End IF
End Function

'for �ഫ xml ���������r�� ex.&
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

'�w��db�J���������S��r��(�S�O�w������ɦW�A�]�������ɦW���i�H�����������Ϊ��A�]���o�˦A�d���˵��ɡA�|�䤣���l�ɮ�
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
