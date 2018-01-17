<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<title>選擇多個項目</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>    
<body style="BACKGROUND-COLOR: #eaf9f5;">
<%
Set conn1 = Server.CreateObject("ADODB.connection")
conn1.Open session("mult") 'Session("ODBCDSN")
Set rsi = Server.CreateObject("ADODB.RecordSet")
%>
<table border="0" align=center cellspacing="0" cellpadding="0">
<tr align=center style="color:blue;font-size:12pt"><td>可選擇的項目</td><td></td><td></td><td></td><td>已選擇的項目</td></tr>
<tr><td>
<SELECT id=select1 name=select1 multiple style="WIDTH: 200px; HEIGHT: 360px">
	<%isql = request("isql")
	rsi.Open isql,conn1,1,1
	while not rsi.EOF%>
		<OPTION value="<%=rsi(0)%>"><%=rsi(1)%></OPTION>
	<%	if not rsi.EOF then rsi.MoveNext
	wend
	rsi.Close%>
</SELECT>
</td>
<td width=10></td>
<td align=center>
<INPUT id=button1 type=button value=">>" name=button1 class="ibuttonr1"><br><br>
<INPUT id=button2 type=button value=">" name=button2 class="ibutton1"><br><br>
<INPUT id=button3 type=button value="<" name=button3 class="ibutton1"><br><br>
<INPUT id=button4 type=button value="<<" name=button4 class="ibuttonr1">
</td>
<td width=10></td>
<td>
<span id=span_scelect2>
<select id=select2 name="select2" multiple style="WIDTH: 200px; HEIGHT: 360px">
</SELECT>
</span>
</table>
<br>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td width="100%" align="center">
		<input type=button name=button5 id=button5 value ="確　定" class="greenbutton" onClick="formSubmit()">
		<input type=button value ="重　填" class="greenbutton" onClick="resetForm()" id=button6 name=button6>
		<input type=button value ="關閉視窗" class="greenbutton" onClick="closeForm()" id=button7 name=button7>
	</td>
	</tr>
</table>
<br>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="font-size:9pt">
	<tr><td>使用方式：請先選擇左邊項目，再點<font color=red>>></font>、<font color=blue>></font>、<font color=blue><</font>、<font color=red><<</font>，即已選擇項目會至右邊清單</tr></td>
	<tr><td><font color=red>>></font>：表選擇左邊所有選擇</tr></td>
	<tr><td><font color=blue>></font>：表將選擇左邊所有反白的項目，新增至已選擇項目</tr></td>
	<tr><td><font color=blue><</font>：表刪除右邊所有反白的項目</tr></td>
	<tr><td><font color=red><<</font>：表刪除右邊所有項目</tr></td>
</table>
</body>
</html>
<script language=vbscript>
sub window_onLoad
	'msgbox "<%=request("value1")%>"
	strhtml = "<SELECT name=select2 id=select2 multiple style='WIDTH: 200px; HEIGHT: 300px'>"
	arvalue1 = split("<%=request("value1")%>",",")
	for k=0 to ubound(arvalue1)-1
		'strhtml = strhtml & "<option value='" & select2(i).value & "'>" & select2(i).text & "</option>"
		for i=0 to select1.length-1
			if select1(i).value=arvalue1(k) then
				strhtml = strhtml & "<option value='" & arvalue1(k) & "'>" & select1(i).text & "</option>"
			end if
		next
	next
	strhtml = strhtml & "</SELECT>"
	document.all.span_scelect2.innerHTML = strhtml
end sub

function formSubmit()
	valueid = ""
	fieldvalue = ""
	for i=0 to select2.length-1
		valueid = valueid & select2(i).value & ","
		fieldvalue = fieldvalue & select2(i).text & "、"
	next
	if fieldvalue<>empty then fieldvalue = left(fieldvalue,len(fieldvalue)-1)
	Execute "window.opener.reg." & "<%=request("field1")%>" & ".value=valueid"
	Execute "window.opener.reg." & "<%=request("field2")%>" & ".value=fieldvalue"
	window.close
end function

function closeForm()
	window.close
end function

function button1_onclick()
	strhtml = "<SELECT name=select2 id=select2 multiple style='WIDTH: 200px; HEIGHT: 300px'>"
	for i=0 to select1.length-1
		strhtml = strhtml & "<option value='" & select1(i).value & "'>" & select1(i).text & "</option>" 		
	next
	strhtml = strhtml & "</SELECT>"
	document.all.span_scelect2.innerHTML = strhtml
end function

function button2_onclick()
	strhtml = "<SELECT name=select2 id=select2 multiple style='WIDTH: 200px; HEIGHT: 300px'>"
	for i=0 to select2.length-1
		strhtml = strhtml & "<option value='" & select2(i).value & "'>" & select2(i).text & "</option>"
	next
	for i=0 to select1.length-1
		if select1(i).selected then
			tflag = false '判斷已有選擇則不再加入,false表尚未選擇該項
			for j=0 to select2.length-1 
				if select1(i).value=select2(j).value then
					tflag = true '表已有選擇該項
					exit for
				end if
			next
			if tflag=false then strhtml = strhtml & "<option value='" & select1(i).value & "'>" & select1(i).text & "</option>"
		end if
	next
	strhtml = strhtml & "</SELECT>"
	document.all.span_scelect2.innerHTML = strhtml
end function

function select1_ondblclick()
	button2_onclick
end function

function button3_onclick()
	strhtml = "<SELECT name=select2 id=select2 multiple style='WIDTH: 200px; HEIGHT: 300px'>"
	for i=0 to select2.length-1
		if select2(i).selected=false then
			strhtml = strhtml & "<option value='" & select2(i).value & "'>" & select2(i).text & "</option>"
		end if
	next
	strhtml = strhtml & "</SELECT>"
	document.all.span_scelect2.innerHTML = strhtml
end function

function select2_ondblclick()
	button3_onclick
end function

function button4_onclick()
	strhtml = "<SELECT name=select2 id=select2 multiple style='WIDTH: 200px; HEIGHT: 300px'>"
	strhtml = strhtml & "</SELECT>"
	document.all.span_scelect2.innerHTML = strhtml
end function

function resetForm
	button4_onclick
end function
</script>
