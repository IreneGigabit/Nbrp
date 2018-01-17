<%
'http://web02/law/inc/xmlgetsqldata.asp
Set cnnf = Server.CreateObject("ADODB.Connection")
Set rsi = Server.CreateObject("ADODB.Recordset")
cnnf.Open session("law")
%>
<?xml version="1.0" encoding="BIG5" standalone="yes"?>
<root>
<%

SQL = trim(Request("SearchSql"))

'SQL = "select agent_na1,agent_na2 from agent where agent_no='2754' and agent_no1='_'"
'SQL = "select seq_area from law where seq_area='LL' and seq=150 and seq1='_'"
Test_Flag = false

if Test_Flag = true then
	Response.Write "<xhead>"
	Response.write "<Found>" & Request("SearchSql") & "</Found>"
	Response.Write "</xhead>"
	Response.Write "</root>"
	Response.End 
end if

if Session("PWD") <> true then
	Response.Write "<xhead>"
	Response.write "<Found>0</Found>"
	Response.Write "</xhead>"
	Response.Write "</root>"
	Response.End 
end if

SQL = replace(Request("SearchSql"),"@","+")

if len(SQL) <= 0 then
	Response.Write "<xhead>"
	Response.write "<Found>0</Found>"
	Response.Write "</xhead>"
	Response.Write "</root>"
	Response.End 
end if
'Response.Write SQL
'Response.End

rsi.Open SQL,cnnf,1,1

if rsi.RecordCount > 0 then
	Response.Write "<xhead>"
	Response.write "<Found>1</Found>"
	For i=0 to rsi.Fields.Count-1
		Response.Write "<" & lcase(rsi(i).Name) & ">" & rsi(i).value & "</" & lcase(rsi(i).Name) & ">"
	Next	
	Response.Write "</xhead>"
else
	Response.Write "<xhead>"
	Response.write "<Found>0</Found>"
	Response.Write "</xhead>"
end if
%>
</root>
<%
set cnnf = nothing
set rsi = nothing
%>
