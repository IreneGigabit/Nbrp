<%@ Language=VBScript%>
<?xml version="1.0" encoding="BIG5"?>
<XMLReply>
<%
crlf = chr(13) & chr(10)
'http://web01/brp/brp2m/xml_returnidname.asp
Set connxml = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
select case request("chkconn")
	case "SYS" connxml.Open Session("ODBCDSN")
	case "P" connxml.Open Session("btbrtdb")
	case "T" connxml.Open Session("btbrtdb")
end select
pcho1 = request("cho1") '條件1
pcho2 = request("cho2") '條件2
pvalue1 = request("value1") '條件傳入的值1
pvalue2 = request("value2") '條件傳入的值2
pname1 = request("name1")	'要傳回的欄位值1
pname2 = request("name2")	'要傳回的欄位值2
ptable = request("tables")	'table name
'pcho1="doc_type"
'pvalue1="A"
'ptable="country"'"ndoc_exp"
'pname1="coun_code"'"doc_code"
'pname2="coun_c"'"doc_c"

	sql = "SELECT " & pname1 & "," & pname2 & " FROM " & ptable  
	if pvalue1<>empty then
		sql = sql & " WHERE " & pcho1 & "='" & pvalue1 & "'"
	end if
	if pvalue2<>empty then
		sql = sql & " and " & pcho2 & "='" & pvalue2 & "'"
	end if
	sql = sql & " ORDER BY 1"
'Response.Write sql
	rs.Open sql,connxml,1,1
	if not rs.EOF then
		while not rs.eof
			Response.Write "<XMLRoot>" & crlf
			Response.Write "<name1>" & trim(rs(0)) & "</name1>" & crlf
			Response.Write "<name2>" & trim(rs(1)) & "</name2>" & crlf
			Response.Write "</XMLRoot>" & crlf
			if not rs.eof then rs.MoveNext
		wend
	end if
	Set rs = Nothing
	set connxml = nothing
%>
</XMLReply>