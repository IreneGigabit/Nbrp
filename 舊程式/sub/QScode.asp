<%
Response.Buffer= true
Response.CacheControl = "no-cache"
Response.Expires =-1
prgid = request("prgid")
copyact=request("copyact")
HTProgCap="�~���d�ߵe��"
HTProgCode=prgid
HTProgPrefix=prgid
Set rsi = Server.CreateObject("ADODB.RecordSet")

'�ݭn�^�g���~�����,�~���m�W���
fscode=trim(request("fscode"))
fscode_name=trim(request("fscode_name"))
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title><%=HTProgCap%></title>
</head>
<!-- #INCLUDE FILE="../inc/server.inc" -->
<!-- #INCLUDE FILE="../calendar/calendar2.inc" -->
<!--#INCLUDE FILE="../sub/Server_conn.vbs" -->
<!-- #INCLUDE FILE="../sub/Server_cbx.vbs" -->
<!-- #INCLUDE FILE="../sub/Client_date.vbs" -->
<!--#INCLUDE FILE="../sub/Client_cbx.vbs" -->
<!--#INCLUDE FILE="../sub/Client_num.vbs" -->
<!--#INCLUDE FILE="../sub/Server_code.vbs" -->
<body>
<center>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td class="FormName">�i<%=HTProgPrefix%>&nbsp;<%=HTProgCap%>�j</td>
		<td width="40%" class="FormName" align="right">
			<!--<a href="imp12Edit.asp?submittask=A&prgid=<%=prgid%>">[�s�W]</a>-->
		</td>
	</tr>
</table>
<hr noshade size="1" color="#000080">
<br>
<Form name="reg" method="POST">
<input type=hidden name="prgid" value="<%=prgid%>">
<input type=hidden name="submittask" value="<%=request("submittask")%>">
<input type=hidden name="copyact" value="<%=copyact%>">
<input type=hidden name="fscode" value="<%=fscode%>">
<input type=hidden name="fscode_name" value="<%=fscode_name%>">
<table border="0" class="bluetable" cellspacing="1" cellpadding="2" width="100%">
	<TR>
		<TD class=lightbluetable align=right>�m�W�G</TD>
		<TD class=whitetablebg align=left>
			<INPUT type="text" name="qsc_name" size="5" maxlength="5">
		</TD>
	</TR>
	<TR>
		<TD class=lightbluetable align=right>�~���G</TD>
		<TD class=whitetablebg>
			<INPUT type="text" name="qscode" size="6" maxlength="5">
		</TD>
	</TR>
</table>
<br>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr class="greenhead">
		<td align=center>
			<%if (HTProgRight AND 2) <> 0 then %>
				<input type="button" value="�d�@��" class="cbutton" onClick="vbscript:formSearchSubmit" id="qrybutton" name="qrybutton">
				<input type="button" value="���@��" class="cbutton" onClick="vbscript:resetForm" id="resbutton" name="resbutton">
				<input type=button class="cbutton" name="btnClose" value ="����">
			<%end if%>
		</td>
	</tr>
</table>
</Form>
</center>
</body>
</html>
<script Language="VBScript">
sub window_onload
end sub
function resetForm()
	reg.reset()
end function
function formSearchSubmit()
	<% if prgid="exp36_6" or prgid="exp6a1" then %>
		if reg.qsc_name.value = "" and reg.qscode.value = "" then
			msgbox "�п�J���@����!"
			reg.qsc_name.focus
			exit function
		end if			
	<% end if %>
	reg.action = "../sub/QScodeList.asp"
    reg.submit()
end function
function btnClose_onclick()
	window.close
end function

</script>
