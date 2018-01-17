<%@ Language=VBScript%>
<%
'入array.mp.agent01501 專利代狂人案件對照暫存檔
'http://web02/brp/sub/array_mp_agent01501.asp -->acc
'http://sin09/brp/sub/array_mp_agent01501.asp -->array
'On Error Resume Next
Session("accmp") = "DSN=arrmp;UID=mmp;PWD=mmp;Database=mp;"	'***Informix-acc-array - mp
session("brpN") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=hjj1823;Initial Catalog=sindbs;Data Source=sin09;"
session("brpC") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=hjj1823;Initial Catalog=sicdbs;Data Source=sic06;"
session("brpS") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=hjj1823;Initial Catalog=sisdbs;Data Source=sis06;"
session("brpK") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=hjj1823;Initial Catalog=sikdbs;Data Source=sik06;"
Set connarr = Server.CreateObject("ADODB.Connection")
connarr.Open Session("accmp")
Set rs1 = Server.CreateObject("ADODB.Recordset") 
Set rs2 = Server.CreateObject("ADODB.Recordset") 

function chknull(p1)
	if trim(p1)<>empty then
		chknull = "'" & p1 & "'"
	else
		chknull = "null"
	end if
end function

getdata "N" 'N:127筆
getdata "C" 'C:69
getdata "S" 'S:42
getdata "K" 'K:54

function getdata(pseq_area)
	cnt = 0
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Session("brp"&pseq_area)

	isql = "select * from agent01501 where seq_area='" & pseq_area & "' order by seq,seq1"
'Response.Write isql & "<BR>"
	rs1.Open isql,connarr
	while not rs1.EOF
		isql = "select agt_no,(select agt_name from agt where agt_no=a.agt_no) as agt_name,apply_no,change_no,end_date" & _
			" from dmp a where seq='"& rs1("seq") &"' and seq1='"& rs1("seq1") &"'"
'Response.Write isql & "<BR>"
		rs2.Open isql,conn,1,1
		if not rs2.EOF then
			if rs2("end_date")<>empty then
				mdydate = string(2-len(month(rs2("end_date"))),"0") & month(rs2("end_date")) & "/" & string(2-len(day(rs2("end_date"))),"0") & day(rs2("end_date")) & "/" & year(rs2("end_date"))
			else
				mdydate = ""
			end if
			usql = "update agent01501 set bragt_no='"& rs2("agt_no") &"',bragt_name='"& rs2("agt_name") &"'," & _
				"brapply_no="& chknull(rs2("apply_no")) &",brchange_no="& chknull(rs2("change_no")) &"," & _
				"end_date="& chknull(mdydate) & _
				" where seq_area='"& pseq_area &"' and seq='"& rs1("seq") &"' and seq1='"& rs1("seq1") &"'"
'Response.Write usql & "<BR>"
			connarr.Execute usql
			cnt = cnt + 1
		end if
		rs2.Close 
		rs1.MoveNext 
	wend
	rs1.Close 
	Response.Write pseq_area & "共update "& cnt &" 筆" &"<BR>"
	set conn = nothing
end function
%>
