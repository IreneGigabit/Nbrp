<!--#include file="class.upload.asp"--><!--�ۭq������-->
<%

response.buffer=true 
mappath_name="\Opt\Opt_File\"


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
   strTestFile = Server.MapPath(strFile)
  
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
      Dim strTestFile,objFSO
      strTestFile = Server.MapPath(strpath)
      set objFSO  = CreateObject("Scripting.FileSystemObject")
	  if chkFileExist(strpath&local_seq2) = 1 then
	     set objFile0= objFSO.GetFile(strTestFile&"\"&local_seq2)
		 objFile0.Delete
	  end if
      set objFile = objFSO.GetFile(strTestFile&"\"&local_seq1)
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
	    file_path="\opt\opt_file\" & folder_name
        call Check_CreateFolder(folder_name)
	    '�^���ɮת����W
		form_name="reg."&session("form_name")
		gsize_name = "reg."&session("size_name")
		gfile_name = "reg."&session("file_name")
		gbtnname = "reg."&session("btnname")
  case else
      Response.redirect "upload_win.asp"
end select

Server.ScriptTimeOut = 1200


'���s�W��
Set mySmartUpload = Server.CreateObject("UpDownExpress.FileUpload")
Set up = New Upload      '�إ�Upload����
Dim strTestFile,objFSO

up.path = file_path      '���w�x�s���|
'cc=
dd=right(up.get_file1("theFile"),len(up.get_file1("theFile"))-InstrRev(up.get_file1("theFile"),"\"))

'Response.Write "aa="&aa&"<br>"
'Response.Write "file_path="&file_path&"<br>"
'Response.Write "dd="&dd&"<br>"
'Response.Write "chk="&chkFileExist(file_path&"\"&dd)&"<br>"
'Response.End
bb=up.get_file(up.get_path("theFile"))
old_file1=right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
if bb = old_file1 then
	a=up.SaveFile("theFile") 
	tfilename=bb
	tsize=up.get_FileSize("theFile")
	aa=file_path&"\"&bb
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
	End IF		
End IF	
%>

<html>
<head>
<title>�W���ɮצ��\!</title>
<script language="vbscript">
    window.opener.<%=form_name%>.value="<%=aa%>"
    if "<%=gsize_name%>" <> empty then
		window.opener.<%=gsize_name%>.value="<%=tsize%>"
	end if
	if "<%=gFile_Name%>" <> empty then	
		window.opener.<%=gfile_name%>.value="<%=tfilename%>"
	end if
	if "<%=gbtnname%>" <> empty then	
		window.opener.<%=gbtnname%>.disabled = true
	end if
    window.close()   
</script>
</head>
<body bgcolor=#ffffff>
</body>
</html>