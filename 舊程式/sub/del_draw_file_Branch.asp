<!--#INCLUDE FILE="../sub/Server_File.asp" -->
<%
response.buffer=true 
'���o (1) gbrFileServerName(�ϩҤW��server name) (2)gFileServerName(��~�ҤW��Server Name)
'(3) gbrDir(�ϩ��ɮת�������|ex.\\sinn01\NPE)    (4)gDir(��~���ɮת�������| ex.\\sin31\FPE_File)
'(3) gbrWebDir(�ϩ��ɮת��������|ex./brp/NPE)    (4)gWebDir(��~���ɮת��������| ex./fexp/FPE_File)
IF session("docbranch")<>empty then
	call getFileServer(session("docbranch"))
	Session("folder_name")="temp/"& session("docbranch") & "/" & session("folder_name")
Else
	call getFileServer("")
End IF	

strtype = Request.QueryString("type")
seq=trim(Request.QueryString("seq"))
draw_file = trim(Request.QueryString("draw_file"))
'file_name = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
file_name = draw_file
btnname = "reg." & trim(request("btnname"))


Dim strTestFile,objFSO
'strTestFile = Server.MapPath(file_path&"/"&file_name)
'�����ϥ�file_name(�ǤJ�������|exp_attach.attach_path)
strTestFile1 = Server.MapPath(file_name)
set objFSO  = CreateObject("Scripting.FileSystemObject")


if chkFileExist_virtual(file_name) = 1 then
	'Response.Write "AAA"
	'�N�R���ɮ׮ɬO�N���word�ƥ��_��(����ɶ����t�@�ӦW�r)
	File_name_new = left(File_name,len(File_name)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (File_name,4)
	renameFile1 File_name,File_name_new
	
	'set objFile= objFSO.GetFile(strTestFile1)
	'objFile.Delete
end if
set objFSO=nothing
%>

<script language="VBScript">	
	if "<%=btnname%>" <> empty then	
		window.opener.<%=btnname%>.disabled = false
	end if
	window.close
</script>
