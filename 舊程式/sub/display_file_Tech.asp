<%
'檢查目錄是否存在
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

'檢查檔案是否存在
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
	Response.Write "  alert ""◎此"&a&"不存在，請確認之後再查!"" " & chr(13)
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
   Response.Write "  alert ""錯誤,請重新查詢!"" " & chr(13)
   Response.Write "  window.close()" & chr(13)
   Response.Write "</SCRIPT>" & chr(13)
   Response.End 
end if

file_path="/brp/Tech_file/"&folder_name

if chkfolderExist(file_path) = 1 then
	if chkfileExist(file_path&"/"&file_name) = 1 then
		call show_photo(file_path&"/"&file_name)
	Else
		call error_msg("圖檔")
	end if
else
	call error_msg("圖檔目錄")
end if

function show_photo(a)
   ImagePath = "http://" & Request.ServerVariables("SERVER_NAME") & a
   'Response.Write ImagePath & "<BR>"
   'Response.End 
   Response.redirect ImagePath
%>
<html>
<head>
<title>商標圖檔資料顯示</title>
</head>
<body onload="vbs:window.focus">
<br><br>
<p><img border="0" src="<%=ImagePath%>" ALT=""></p>
</body>
</html>
<%end function%>



