
<html>
<head>
<title><%=cont%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<!--#include file="class.upload.asp"--><!--�ۭq������-->
<%
response.buffer=true 
mappath_name="\brp\Tech_file\"

'�ˬd�ؿ��O�_�s�b
function chkfolderExist(strFolder)
	Dim strTestFolder,objFSO
	'strTestFolder = Server.MapPath(strFolder)
	strTestFolder = strFolder
	'Response.Write "strTestFolder=" & strTestFolder & "<br>"
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
	'strTestFile = Server.MapPath(strFile)
	 strTestFile =  strFile
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTestFile) then
		chkFileExist=1
	else
		chkFileExist=0
	end if
	set objFSO=nothing
end function

'�إߥؿ�
function Check_CreateFolder(strFolder)
	Dim strTestFolder,objFSO
	Dim aryfolder,i

	set objFSO = CreateObject("Scripting.FileSystemObject")
	strTestFolder = Server.MapPath(mappath_name)
	aryfolder=split(strfolder,"/")
	for i=0 to ubound(aryfolder)
		strTestFolder = strTestFolder & "\" & aryfolder(i)
		'Response.Write strTestFolder&"<br>"
		'Response.End
		
		if chkfolderExist(strTestFolder) = 1 then
		else
			objFSO.CreateFolder(strTestFolder)
		end if	
	next 
	set objFSO=nothing
end function

'�\���ɮ�
function CoverFile(strpath,local_seq1,local_seq2)
	Dim strTestFile,objFSO
	strTestFile = Server.MapPath(strpath)
	set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.MoveFile strTestFile&"\"&local_seq2,strTestFile&"\"&local_seq2&"_tmp"
	set objFSO=nothing
	call renameFile(strpath,local_seq1,local_seq2)
	call DelFile(strpath,local_seq1,local_seq2)
end function

'�屼�ɮ�
function DelFile(strpath,local_seq1,local_seq2)
	Dim strTestFile,objFSO
	strTestFile = Server.MapPath(strpath)
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTestFile&"\"&local_seq2&"_tmp") = True then
		objFSO.DeleteFile strTestFile&"\"&local_seq2&"_tmp"
	end if 
	set objFSO=nothing
end function
'�屼�ɮ�1
function DelFile1(strpath,local_seq2)
	Dim strTestFile,objFSO
	strTestFile = Server.MapPath(strpath)
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTestFile&"\"&local_seq2) = True then
		objFSO.DeleteFile strTestFile&"\"&local_seq2
	end if 
	set objFSO=nothing
end function

'�ɮ׭��s�R�W
function renameFile(strpath,local_seq1,local_seq2)
	'Response.Write "strpath="&strpath&"<br>"
	'Response.Write "local_seq1="&local_seq1&"<br>"
	'Response.Write "local_seq2="&local_seq2&"<br>"
	'Response.End
	
	Dim strTestFile,objFSO
	strTestFile = Server.MapPath(strpath)
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	if chkFileExist(strpath&local_seq2) = 1 then
		set objFile0= objFSO.GetFile(strTestFile&"\"&local_seq2)
		objFile0.Delete
	end if
	set objFile = objFSO.GetFile(strTestFile&"/"&local_seq1)
	objFile.Move(strTestFile&"\"&local_seq2)
	set objFSO=nothing
end function


function MoveFile(strpath,local_seq1,local_seq2)
	if chkFileExist(strpath&local_seq2) = 1 then
		Response.Write "<SCRIPT LANGUAGE=vbs>"& chr(13)
		Response.Write " alert(""���ɮפw�s�b! �w�л\�ɮ�."")"& chr(13)
		Response.Write "</SCRIPT>"& chr(13)
		call CoverFile(strpath,local_seq1,local_seq2)
	else
		call renameFile(strpath,local_seq1,local_seq2)
	end if
end function

Dim mySmartUpload

select case lcase(session("type"))
	case "doc"
		seq = session("seq")
		folder_name = session("folder_name")
		prefix_name = session("prefix_name")
		draw_file = session("draw_file")
		old_file = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
		file_path="\brp\Tech_file\" & folder_name
        call Check_CreateFolder(folder_name)
	    '�^���ɮת����W
		form_name="reg."&session("form_name")
		gsize_name = "reg."&session("size_name")
		gfile_name = "reg."&session("file_name")
		gbtnname = "reg."&session("btnname")
		doc_in_scode = "reg."&session("doc_in_scode")
		doc_in_date = "reg."&session("doc_in_date")
	case else
		Response.redirect "upload_win.asp"
end select

Server.ScriptTimeOut = 1200

'���s�W��
Set mySmartUpload = Server.CreateObject("UpDownExpress.FileUpload")
Set up = New Upload      '�إ�Upload����
Dim strTestFile,objFSO

Response.Write file_path &"<BR>"
'up.path = replace(file_path,"/","\")      '���w�x�s���|

up.path = file_path      '���w�x�s���|
dd=right(up.get_file1("theFile"),len(up.get_file1("theFile"))-InstrRev(up.get_file1("theFile"),"\"))
bb=up.get_file(up.get_path("theFile"))
'old_file1=up.get_file(up.get_path("theFile"))
old_file1=right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))

if bb = old_file1 then
	a=up.SaveFile("theFile") 
	tfilename=bb
	tsize=up.get_FileSize("theFile")
	aa=file_path&"\"&bb
	call renameFile(file_path,up.get_file1("theFile"),ee)
	'�ˬd�W�Ǫ��ɮ׬O�_�s�b
	Response.Write "<SCRIPT LANGUAGE=vbs>"& chr(13)
	Response.Write " alert(""���ɮפw�s�b! �w�л\�ɮ�."")"& chr(13)
	Response.Write "</SCRIPT>"& chr(13)
else
	if chkFileExist(file_path&"\"&dd) = 1 then%>
		<SCRIPT LANGUAGE=vbs>
		msgbox "���ɮפw�g�s�b!!" & chr(10) & chr(10) & "�бN���ɮק�W�A�í��s�W��!!" 
		window.close
		</SCRIPT>
	<%Else
		a=up.SaveFile("theFile") 
		tfilename=bb
		'Response.Write tfilename&"<br>"
		'Response.End
		
		tsize=up.get_FileSize("theFile")
		aa=file_path&"\"&ee
		call renameFile(file_path,up.get_file(up.get_path("theFile")),ee)
	End IF		
End IF	
%>
<script language="vbscript">
'msgbox "<%=form_name%>"
    window.opener.<%=form_name%>.value="<%=aa%>"
    if "<%=gsize_name%>" <> empty then
		window.opener.<%=gsize_name%>.value="<%=tsize%>"
	end if
	if "<%=gFile_Name%>" <> empty then	
		window.opener.<%=gfile_name%>.value="<%=ee%>"
	end if
	if "<%=gbtnname%>" <> empty then	
		window.opener.<%=gbtnname%>.disabled = true
	end if
	IF "<%=doc_in_date%>" <> empty then
		window.opener.<%=doc_in_date%>.value="<%=date()%>"
	End IF
	IF "<%=doc_in_scode%>" <> empty then
		window.opener.<%=doc_in_scode%>.value="<%=session("scode")%>"
	End IF
    window.close()   
</script>
</head>
<body bgcolor=#ffffff>
</body>
</html>