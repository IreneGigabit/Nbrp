<%
response.buffer=true 
gdept = "P"

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

strtype = Request.QueryString("type")
seq=trim(Request.QueryString("seq"))
draw_file = trim(Request.QueryString("draw_file"))
file_name = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
folder_name = request("folder_name")
cust_area=left(Request("cust_area"),1)&gdept
file_path="/brp/Tech_file/"&folder_name
btnname = "reg." & trim(request("btnname"))


Dim strTestFile,objFSO
strTestFile = Server.MapPath(file_path&"/"&file_name)
set objFSO  = CreateObject("Scripting.FileSystemObject")
if chkFileExist(file_path&"/"&file_name) = 1 then
	set objFile= objFSO.GetFile(strTestFile)
	objFile.Delete
end if
set objFSO=nothing
%>

<script language="VBScript">	
	if "<%=btnname%>" <> empty then	
		window.opener.<%=btnname%>.disabled = false
	end if
	window.close
</script>
