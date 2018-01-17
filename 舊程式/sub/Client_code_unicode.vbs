<script Language="vbScript">
' 顯示Unicode字元(避免出現?)
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
	End iF	
End Function

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

</script>
