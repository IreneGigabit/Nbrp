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
			<a href="multi_Qscode.asp?prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>">[�^�d�ߵe��]</a>
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
			'Response.Write totRec
			'Response.end
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
				|<a href="multi_QScodeList.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%><%=tlink%>&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%>"><font size="2">�W�@��</font></a>
			<%end if%>
			<% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage) < RSreg.PageCount  then %>
				|<a href="multi_QScodeList.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%><%=tlink%>&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%>"><font size="2">�U�@��</font></a>
			<%end if%></font>
			</td>
		</tr>
	</table>
<input type=hidden name="prgid" value="<%=prgid%>">
<table border="0" class="bluetable" cellspacing="1" cellpadding="2" width="100%">
	<TR>
		<%if copyact = "1" then%>
		<TD class=lightbluetable align=center><a href="vbscript:selectall()">����</a></TD>
		<%end if%>
		<TD class=lightbluetable align=center>���u�m�W</TD>
		<TD class=lightbluetable align=center>���u�~��</TD>
	</TR>
	<%'While Not RSreg.EOF
	for i=1 to PerPageSize
	%>
		<tr align="center" class="sfont9">
			<input type="hidden" name="scode<%=i%>" value="<%=RSreg("scode")%>">
			<%if copyact = "1" then%>
			<td nowrap>
				<!--<td style="cursor: hand;background-color:white" onmouseover="vbs:me.style.color='red'" onmouseout="vbs:me.style.color='black'" nowrap onclick="VBScript:ScodeClick '<%=RSreg("scode")%>','<%=RSreg("sc_name")%>'">-->
				<input type="checkbox" id=chkflag<%=i%> name=chkflag<%=i%> onclick="vbscript:chkflag_onclick <%=i%>">
				<input type="hidden" id=hchkflag<%=i%> name=hchkflag<%=i%>>
			</td>
			<%end if%>
			<td nowrap><%=RSreg("sc_name")%></td>			
			<td nowrap><%=RSreg("scode")%></td>
		</tr>
	<%
		RSreg.MoveNext
		if RSreg.EOF then 
			i = i + 1
			exit for
		end if
	'Wend
	next
	%>
	<input type="hidden" id="chknum" name="chknum" value=<%=i-1%>>
</table>
	<%if copyact = "1" then%>
		<br>
		<input type="button" value="�a�^�Ŀ���" class="cbutton" onclick="VBScript:ScodeClick 'N'" id="qrybutton" name="qrybutton">&nbsp;&nbsp;
		<input type="button" value="�a�^�Ŀ��ơ���������" class="cbutton" onclick="VBScript:ScodeClick 'Y'" id="qrybutton" name="qrybutton">&nbsp;&nbsp;
		<input type="button" value="��������" class="cbutton" onclick="VBScript:window.close" id="qrybutton" name="qrybutton">
	<%end if%>		
<!--<p><font color=blue>*** ���I�糧�ҽs���N��Ʊa�^���o��@�~ ***</font></p>-->
<%else%>
	<div align="center"><font color="red" size=2>=== �d�L���===</font></div>
<%end if%></Form>
</center>
</body>
</html>
<br>
<%if copyact = "1" then%>
	<!--<font color="red">*</font> �I���~���A�i�N�~���a�^!!<br>-->
<%end if%>
<script Language="VBScript">
sub window_onload
	'window.opener.close
end sub
sub GoPage_OnChange()               '���ܭ���
	on error resume next
	newPage=reg.GoPage.value     
	flag = true
	window.location.href="multi_QScodeList.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%><%=tlink%>"	
end sub      
sub PerPage_OnChange()              '���w�C������  
	on error resume next
	newPerPage=reg.PerPage.value
	window.location.href="multi_QScodeList.asp?nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage & "&prgid=<%=prgid%>&copyact=<%=copyact%>&fscode=<%=fscode%>&fscode_name=<%=fscode_name%><%=tlink%>"
end sub

'�T�w��N��Ʊa�^�B�z
function ScodeClick(paction)
Dim i,old_list,select_list,new_list

	execute "old_list=window.opener.reg." & "<%=fscode%>" & ".value"
	
	'msgbox "old_list=" & old_list
	'�ˬd�O�_���Ŀ�
	totnum =0
	for i = 1 to reg.chknum.value
		execute "set tchkflag=reg.hchkflag" & i 
		if tchkflag.value = "Y" then
			'���~���r��a�^�A�p�G���ƪ��N���n�[�J
			if instr(old_list,eval("reg.scode" & i&".value")) = 0 then
				select_list = select_list & trim(eval("reg.scode" & i&".value")) & ";"
				'msgbox i & ":" & trim(eval("reg.scode" & i&".value"))
			end if		
			totnum = totnum + 1								
		end if
	next
	
	'msgbox totnum
	if totnum = 0 then
		msgbox "�ܤֻݤĿ�@�����!!"
		exit function
	end if	
	
	if trim(old_list) <> empty then
		if right(old_list,1) = ";" then
			new_list = old_list & select_list
		else
			new_list = old_list & ";" & select_list
		end if
	else
		new_list = select_list
	end if
	
	if trim(new_list) <> empty then
		if right(new_list,1) = ";" then	
			new_list = mid(new_list,1,len(new_list)-1)
		end if
	end if
'msgbox "new_list=" & new_list
	execute "window.opener.reg." & "<%=fscode%>" & ".value =new_list"
	'execute "window.opener.reg." & "<%=fscode_name%>" & ".value = x2"
	if paction = "Y" then
		window.close
	end if
end function

'����\��
function selectall()
	for i = 1 to reg.chknum.value
		execute "reg.chkflag" & i & ".checked=true"
		chkflag_onclick(i)
	next
end function

function chkflag_onclick(pchknum)
	tstr1="Y" 
	tstr2="N" 
	if eval("reg.chkflag"& pchknum & ".checked") then
		execute "reg.hchkflag" & pchknum & ".value=tstr1"
	else
		execute "reg.hchkflag" & pchknum & ".value=tstr2"		
	end if
End function
</script>