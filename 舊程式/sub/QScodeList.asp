<%
Response.Buffer= true
Response.CacheControl = "no-cache"
Response.Expires =-1
prgid = request("prgid")
copyact = request("copyact")

HTProgCap="�~���d�ߵ��G�e��"
HTProgCode=prgid
HTProgPrefix=prgid
'Response.Write prgid
qsc_name=request("qsc_name")
qscode=request("qscode")

'�ݭn�^�g���~�����,�~���m�W���
fscode=trim(request("fscode"))
fscode_name=trim(request("fscode_name"))
'Response.Write "1=" & fscode & "<br>"
'Response.Write "2=" & fscode_name & "<br>"
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
<body>
<%
Set RSreg = Server.CreateObject("ADODB.RecordSet")
Set rsi = Server.CreateObject("ADODB.RecordSet")

isql = "select * from sysctrl.dbo.scode where (end_date is null or end_date >= '" & date() & "') "
if qsc_name <> empty then
	isql = isql & " and sc_name like '%" & qsc_name & "%'"
	tlink = tlink & "&qsc_name=" & qsc_name
end if
if qscode <> empty then
	isql = isql & " and scode like '%" & qscode & "%'"
	tlink = tlink & "&qscode=" & qscode
end if

isql = isql & " order by sscode "
'Response.Write isql & "<br>"
'Response.End
%>
<center>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td class="FormName">�i<%=HTProgPrefix%>&nbsp;<%=HTProgCap%>�j</td>
		<td width="40%" class="FormName" align="right">
			<a href="vbscript:window.history.back">[�^�d�ߵe��]</a>
			<a href="vbscript:window.close()">[��������]</a>
		</td>
	</tr>
</table>
<hr noshade size="1" color="#000080">
<Form name="reg" method="POST">
	<%
	RSreg.Open isql,conn,1,1
	if RSreg.EOF then
		Response.Write "��Ƥ��s�b, �Э��s�d��!!&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.End
	end if
	if not RSreg.EOF then
		nowPage=Request.QueryString("nowPage")  '�{�b����
		if Not RSreg.eof then
			totRec = RSreg.Recordcount       '�`����
			if totRec>0 then 
				PerPageSize = 10            '�C������
				if PerPageSize <= 0 then PerPageSize=10
				RSreg.PageSize=PerPageSize
				if cint(nowPage)<1 then 
					nowPage=1
				elseif cint(nowPage) > RSreg.PageCount then 
					nowPage=RSreg.PageCount
				end if
				Session("QueryPage_No") = nowPage
				RSreg.AbsolutePage = nowPage
				totPage=RSreg.PageCount       '�`����
			end if
		end if
	%>
	<table name="Page" width="100%" cellspacing="1" cellpadding="0" border="0">
		<tr align="center">
			<td><input type="hidden" name="agrs" value="<%=agrs%>"></td>
			<td align="center" colspan="6">
			<font color="rgb(0,64,128)" size=2> ��
			<font color="#FF0000" size=2><%=nowPage%>/<%=totPage%></font> �� | ��Ʀ@
			<font color="#FF0000" size=2><%=totRec%></font>
			<font color="rgb(0,64,128)" size=2>�� | ���ܲ�
			<select id="GoPage" name="GoPage" onchange="vbscript:GoPage_OnChange" size="1" style="color:#000000">
				<%For iPage=1 to totPage%>
					<option value="<%=iPage%>" <%if iPage=cint(nowPage) then Response.Write "selected style='color:red'" end if%>><%=iPage%></option>
				<%Next%>
			</select>
			��</font>
			<%if cint(nowPage) <>1 then %>
				|<a href="QScodeList.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%><%=tlink%>&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%>"><font size="2">�W�@��</font></a>
			<%end if%>
			<% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage) < RSreg.PageCount  then %>
				|<a href="QScodeList.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%><%=tlink%>&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%>"><font size="2">�U�@��</font></a>
			<%end if%></font>
			</td>
		</tr>
	</table>
<input type=hidden name="prgid" value="<%=prgid%>">
<table border="0" class="bluetable" cellspacing="1" cellpadding="2" width="100%">
	<TR>
		<TD class=lightbluetable align=center>���u�~��</TD>
		<TD class=lightbluetable align=center>���u�m�W</TD>
	</TR>
	<%'While Not RSreg.EOF
	for i=1 to PerPageSize
	%>
		<tr align="center" class="sfont9">
			<%if copyact = "1" then%>
				<td style="cursor: hand;background-color:white" onmouseover="vbs:me.style.color='red'" onmouseout="vbs:me.style.color='black'" nowrap onclick="VBScript:ScodeClick '<%=RSreg("scode")%>','<%=RSreg("sc_name")%>'">
			<%else%>
				<td nowrap>
			<%end if%>
				<%=RSreg("scode")%>
			</td>
			<td nowrap><%=RSreg("sc_name")%></td>
		</tr>
	<%
		RSreg.MoveNext
		if RSreg.EOF then exit for
	'Wend
	next
	%>
</table>
<!--<p><font color=blue>*** ���I�糧�ҽs���N��Ʊa�^���o��@�~ ***</font></p>-->
<%else%>
	<div align="center"><font color="red" size=2>=== �d�L�ץ���===</font></div>
<%end if%></Form>
</center>
</body>
</html>
<br>
<%if copyact = "1" then%>
	<font color="red">*</font> �I���~���A�i�N�~���a�^!!<br>
<%end if%>
<script Language="VBScript">
sub window_onload
	'window.opener.close
end sub
sub GoPage_OnChange()               '���ܭ���
	on error resume next
	newPage=reg.GoPage.value     
	flag = true
	window.location.href="QScodeList.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%><%=tlink%>"	
end sub      
sub PerPage_OnChange()              '���w�C������  
	on error resume next
	newPerPage=reg.PerPage.value
	window.location.href="QScodeList.asp?nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage & "&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%><%=tlink%>"
end sub
function ScodeClick(x1,x2)
	execute "window.opener.reg." & "<%=fscode%>" & ".value = x1"
	execute "window.opener.reg." & "<%=fscode_name%>" & ".value = x2"
	window.close
end function
</script>