<%@ Language=VBScript %>
<%
Set conn1 = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
conn1.Open Session("ODBCDSN")
sql = "insert into rec_log values(null,'" & ptableid & "','" & pprgid & "','" & pscode & "','" & date() & "','" & pnote & "')"
'Response.Write SQL & "<BR>"
'Response.End
rs.Open sql,conn1,1,1,adcmdtext
set rs = nothing
set conn1 = nothing
%>
<script language="VBScript">
'	window.opener.reg.btnseq1.value = "<%=treturn%>"
	window.close
</script>
