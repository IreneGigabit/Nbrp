<%
'�ˬd�ؿ��O�_�s�b
function chkfolderExist(strFolder)
	Dim strTestFolder,objFSO
	strTestFolder = Server.MapPath(strFolder)
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FolderExists(strTestFolder) then
		chkfolderExist=1
	else
		chkfolderExist=0
	end if
	set objFSO=nothing
end function

'�ˬd�ɮ׬O�_�s�b
function chkFileExist(strFile)
	Dim strTestFile,objFSO
	strTestFile = Server.MapPath(strFile)
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTestFile) then
		chkFileExist=1
	else
		chkFileExist=0
	end if
	set objFSO=nothing
end function

function error_msg(a)
	Response.Write "<SCRIPT LANGUAGE=vbscript>" & chr(13)
	Response.Write "  alert ""����"&a&"���s�b�A�нT�{����A�d!"" " & chr(13)
	Response.Write "  window.close()" & chr(13)
	Response.Write "</SCRIPT>" & chr(13)
	Response.End 
end function

strtype = Request.QueryString("type")
seq=trim(Request.QueryString("seq"))
draw_file = trim(Request.QueryString("draw_file"))
file_name = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
folder_name = request("folder_name")
 
if isnull(strtype) or strtype = "" or strtype <> "doc" then
   Response.Write "<SCRIPT LANGUAGE=vbscript>" & chr(13)
   Response.Write "  alert ""���~,�Э��s�d��!"" " & chr(13)
   Response.Write "  window.close()" & chr(13)
   Response.Write "</SCRIPT>" & chr(13)
   Response.End 
end if
if left(session("prgid"),3) = "brp" or session("sendprgid") = "TRANDMP" or request("dept_area")="P" then
    if request("databr_branch")<>empty then
        file_path="/brp/"&request("databr_branch")&"P/"&folder_name
    else
	    file_path="/brp/"&session("se_branch")&"P/"&folder_name
	end if
else
    if request("databr_branch")<>empty then
    	file_path="/brp/"&request("databr_branch")&"PE/"&folder_name
    else
    	file_path="/brp/"&session("se_branch")&"PE/"&folder_name
	end if
end if
if session("scode")="m983" then
'    Response.Write  file_path&"/"&file_name& "<BR>"
'    Response.End 
end if

if chkfolderExist(file_path) = 1 then
	if chkfileExist(file_path&"/"&file_name) = 1 then
		call show_photo(file_path&"/"&file_name)
	Else
		call error_msg("����")
	end if
else
	call error_msg("���ɥؿ�")
end if

function show_photo(a)
   ImagePath = "http://" & Request.ServerVariables("SERVER_NAME") & a
   'Response.Write ImagePath & "<BR>"
   'Response.End 
   Response.redirect ImagePath
%>
<html>
<head>
<title>�Ӽй��ɸ�����</title>
</head>
<body onload="vbs:window.focus">
<br><br>
<p><img border="0" src="<%=ImagePath%>" ALT=""></p>
</body>
</html>
<%end function%>



