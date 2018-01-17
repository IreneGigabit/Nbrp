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
'response.write session("btbrtdb")
'response.end
conn.Open session("btbrtdb")
Set cnn = Server.CreateObject("ADODB.Connection")
if lcase(left(Request.ServerVariables("SERVER_NAME"),3))="sinn" then
	if lcase(HTProgCode)="si00w01" or lcase(HTProgCode)="si00w02" or lcase(HTProgCode)="si00w03" _
	or lcase(HTProgCode)="si00w04" or lcase(HTProgCode)="si00w11" or lcase(HTProgCode)="si00w12" _
	or lcase(HTProgCode)="si00w24" or lcase(HTProgCode)="si00w25" _
	or lcase(HTProgCode)="ap1" or lcase(HTProgCode)="ap2" or lcase(HTProgCode)="ap3" _
	or lcase(HTProgCode)="ap4" or lcase(HTProgCode)="ap5" then
		cnn.Open session("ODBCDSNsin09")
	else
		cnn.Open session("ODBCDSN")
	end if
else
	cnn.Open session("ODBCDSN")
end if

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
		'Response.Write SQL & "<br>"
		'Response.End 
 '  Sql = "SELECT rights FROM GrpAP INNER JOIN scode ON GrpAP.GrpID IN (" & IDStr & ")" & _
 '        "WHERE GrpAP.beg_date<=getdate() and GrpAP.end_date>=getdate()" & _
 '        " and scode.scode='" & session("scode") & "' and GrpAP.APcode='" & HTProgCode & "'"
 'response.write SQL
   set RS=cnn.execute (Sql)
if rs.eof then %>
<script language=VBScript>
	Alert ("無使用權限！")
	top.location.href="../default.asp"
</script>
<%
else
	if rs("rights")=0 then%>
		<script language=VBScript>
			Alert ("無使用權限！")
			if brp then
				top.location.href= "http://" & location.hostname & "brp/default.asp"
			else
				top.location.href="../default.asp"
			end if
		</script>
	<%else

		Sql = "SELECT APnameC FROM AP Where APCode = '"& HTProgCode &"' AND SYScode='" & xCode & "'"
		set RScom=cnn.execute (Sql)
		If Not rscom.EOF Then
			Title = rscom("APnameC") & "&nbsp;管理"
			Title2 = rscom("APnameC")   
		end if
	end if   

	HTProgRight=rs("rights")
end if

%>
<%
'20170215 Response Sql for Debug
Function showLog(text)
	'if session("LoginGrp")="BTBRTADMIN" then
	if session("se_scode")="m1583" then
		Response.Write(text & "<hr>")
	end if
End Function
Function showLogEnd(text)
	'if session("LoginGrp")="BTBRTADMIN" then
	if session("se_scode")="m1583" then
		Response.Write(text & "<hr>")
		Response.end()
	end if
End Function
%>