<%
If Not Session("pwd") Then 
%>
<script language=VBScript>
	Alert ("系統停滯時間逾時！請重新登入！")
	top.location.href="../default.asp"
</script>
<%  
End If %> 
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open session("btbrtdb")
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.Open session("ODBCDSN")

progPath = request.serverVariables("PATH_INFO")
posPath = InStrRev(progPath, "/")
if posPath > 0 then		progPath = mid(progPath, posPath+1)
posPath = Instr(progPath,".")
if posPath > 0 then		progpath = left(progPath, posPath-1)
	xCode = session("Syscode")
	xPos=Instr(session("LoginGrp"),",")
	if xPos>0 then
		IDStr=replace(session("LoginGrp"),", ","','")
		IDStr="'"&IDStr&"'"
	else
		IDStr="'" & session("LoginGrp") & "'"
	end if
		SQL = "SELECT Rights FROM LoginAP WHERE " & _
					"SYScode = '" & xCode & "' AND LoginGrp = " & IDStr & " AND " & _
					"beg_date <= GETDATE() AND end_date >= GETDATE() AND " & _
					"APcode = '" & HTProgCode & "'"  
 '  Sql = "SELECT rights FROM GrpAP INNER JOIN scode ON GrpAP.GrpID IN (" & IDStr & ")" & _
 '        "WHERE GrpAP.beg_date<=getdate() and GrpAP.end_date>=getdate()" & _
 '        " and scode.scode='" & session("scode") & "' and GrpAP.APcode='" & HTProgCode & "'"
 'response.write sql &"<BR>"
 'response.end
   set RS=cnn.execute (Sql)
if rs.eof then %>
<script language=VBScript>
	Alert ("無使用權限！")
	top.location.href="../default.asp"
</script>
<%
else

   Sql = "SELECT APnameC FROM AP Where APCode = '"& HTProgCode &"' AND SYScode='" & xCode & "'"
   set RScom=cnn.execute (Sql)
   If Not rscom.EOF Then
   	Title = rscom("APnameC") & "&nbsp;管理"
   	Title2 = rscom("APnameC")   
   end if
   

   HTProgRight=rs("rights")
end if

if ((HTProgRight AND 1) = 0) AND (lCase(Right(progPath,5))="query") then
%>
<HTML>
<BODY>
<SCRIPT LANGUAGE=vbs>
msgBox "‧該系統未授權"
window.history.back
</SCRIPT>
</BODY>
</HTML>
<%  
Response.End 
End If 
%> 
