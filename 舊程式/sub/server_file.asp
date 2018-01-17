<%
'�w�q�@�ΰѼ�
gdept = "P"
gbrfileservername = ""
gbrfilesmapervername = ""  '����ɮץD��
gfileservername = ""
gbrDir = ""
gDir = ""
gbrWebDir = ""
gWebDir = ""
gbrDbDir = ""
gcustDbDir = ""

'���oserver name
'���o (1) gbrFileServerName(�ϩҤW��server name) (2)gFileServerName(��~�ҤW��Server Name)
'(3) gbrDir(�ϩ��ɮת�������|ex.\\sinn03\NPE)    (4)gDir(��~���ɮת�������| ex.\\sin31\FPE_File)
'(5) gbrWebDir(�ϩ��ɮת��������|ex./brp/NPE)    (6)gWebDir(��~���ɮת��������| ex./fexp/FPE_File)
'(7) gbrDbDir(�ϩҽдڳ檺�������|ex./brp/brdb_file) (8) gcustDbDir(�ϩҹ�ʱb�Ȩ窺�������|ex./brp/custdb_file)
function getFileServer(pbr_branch)
	Select Case Request.ServerVariables("SERVER_NAME")
		Case "web01"
			gbrfileservername = "web01"
			gfileservername = "web01"
			gbrfilesmapervername = "web01"
		Case "web02"
			gbrfileservername = "web02"
			gfileservername = "web02"
			gbrfilesmapervername = "web02"
		Case "bik02"
			gbrfileservername = "bik02"
			gfileservername = "bik02"
			gbrfilesmapervername = "bik02"
		Case else
			Select Case pbr_branch
			Case "N" 
				gbrfileservername = "sinn03"
				gbrfilesmapervername = "sinn01"  '����ɮץD��
			Case "C" 
				gbrfileservername = "sic09"
				gbrfilesmapervername = "sic08"  '����ɮץD��
			Case "S" 
				gbrfileservername = "sis09"
				gbrfilesmapervername = "sis08"   '����ɮץD��
			Case "K"
				gbrfileservername = "sik09"
				gbrfilesmapervername = "sik08"   '����ɮץD��
			End Select 
			gfileservername = "sin31"	'?�W�u�ɻݭק�
	end Select
	gbrDir = "\\" & gbrfileservername & "\" & pbr_branch & "PE"
	gDir =  "\\" & gfileservername & "\" & "FPE_file"
	
	'Response.Write session("prgid") & "<BR>"
	if left(session("prgid"),3) = "brp" or session("sendprgid") = "TRANDMP" then
		gbrWebDir = "/" & "brp" & "/" & pbr_branch & "P"
		gWebDir =  "/" & "mg" & "/" & "FPE_File"
	else
		gbrWebDir = "/" & "brp" & "/" & pbr_branch & "PE"
		gWebDir =  "/" & "Fexp" & "/" & "FPE_File"
	end if
	'2012/5/2�W�[�дڳ�ι�ʱb�Ȩ�
	gbrDbDir = "/brp/brdb_file"
	gcustDbDir = "/brp/custdb_file"
	'Response.Write "gbrWebDir=" & gbrWebDir & "<br>"
	'Response.Write "gWebDir=" & gWebDir & "<br>"
	'Response.Write "gbrDir=" & gbrDir & "<br>"
	'Response.Write "gDir=" & gDir & "<br>"
	'Response.end
end function

'�ˬd�ϺЬO�_�s�b
function chkDriveExist(strFolder)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	strTestFolder = strFolder
	if objFSO.DriveExists(strTestFolder) then
		chkDriveExist=1
	else
		chkDriveExist=0
	end if
	set objFSO = nothing
end function

'�ˬd�ؿ��O�_�s�b
function chkfolderExist(strFolder)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	strTestFolder = strFolder
	if objFSO.FolderExists(strTestFolder) then
		chkfolderExist=1
	else
		chkfolderExist=0
	end if
	set objFSO = nothing
end function

'�إߥؿ�(�ϥ� share folder�覡,�����Q�ι�����| create folder)
'ex.pfldr1=\\sinn03\NPE  pfldr2=_/174/17432
function Check_CreateFolder(pfldr1,pfldr2)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim aryfolder,i
	strTestFolder = ""
	strFolder = ""
	
	aryfolder=split(pfldr2,"/")
	for i=0 to ubound(aryfolder)
		strTestFolder = pfldr1 & "\" & aryfolder(i)
		strFolder = strFolder &"\" & aryfolder(i)
		strTestFolder = pfldr1 & strFolder
		'Response.Write "strTestFolder="&strTestFolder&"<br>"
		'Response.Write "chkfolderExist="&chkfolderExist(strTestFolder)&"<br>"
		'rMrk = rMrk + "#@#" + strTestFolder
        'rMrk = rMrk & "#@#aaa-"& err.number &"---"& err.Description 
		'Response.end
		if chkfolderExist(strTestFolder) = 1 then
		else
			'Response.Write strTestFolder&"<br>"
			'Response.end
            'rMrk = rMrk & "#@#bbb-"& err.number &"---"& err.Description 
			objFSO.CreateFolder(strTestFolder)
            'rMrk = rMrk & "#@#ccc-"& err.number &"---"& err.Description 
		end if			
	next 
	'Response.End 
	set objFSO = nothing
end function

'�إߥؿ�(�ǤJvirtual directory,�A��mappath�覡�ഫ��������|,Create folder )
function Check_CreateFolder_virtual(strsite,strFolder)
	Dim strTestFolder,objFSO
	Dim aryfolder,i

	set objFSO = CreateObject("Scripting.FileSystemObject")
	strTestFolder = Server.MapPath(strsite)
	'response.Write strsite &"<BR>"
	'response.end
	
	aryfolder=split(strfolder,"/")
	for i=0 to ubound(aryfolder)
		strTestFolder = strTestFolder & "\" & aryfolder(i)
		'response.Write strTestFolder &"<BR>"
		if chkfolderExist(strTestFolder) = 1 then
		else
			objFSO.CreateFolder(strTestFolder)
		end if	
	next 
	set objFSO=nothing
	'response.End 
end function

'�ˬd�ɮ׬O�_�s�b(�ǤJ����������|,ex.share folder�覡 : //web02/FPE_File)
function chkFileExist(strFile)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strFile) then
		chkFileExist=1
	else
		chkFileExist=0
	end if
	set objFSO = nothing
end function

'�ˬd�ɮ׬O�_�s�b(�ǤJ�������ؿ� ex./fexp/FPE_File/_/700/70002)
function chkFileExist_virtual(strFile)
	Dim strTestFolder,objFSO
	if session("scode")="m983" then
		'Response.Write "===========chkFileExist_virtual========begin" & "<BR>"
		'Response.Write "chkFileExist_virtual="& strFile & "<BR>"
		'response.End 
	end if
	set objFSO = CreateObject("Scripting.FileSystemObject")
	strTestFile = Server.MapPath(strFile)
	'if session("scode")="admin" then
	'	Response.Write "chkFileExist_virtual="& strTestFile & "<BR>"
	'end if
	if objFSO.FileExists(strTestFile) then
		chkFileExist_virtual=1
	'	if session("scode")="admin" then
	'		Response.Write "chkFileExist_virtual=1<BR>"
	'	end if
	else
		chkFileExist_virtual=0
	'	if session("scode")="admin" then
	'		Response.Write "chkFileExist_virtual=0<BR>"
	'	end if
	end if
	set objFSO = nothing
	'if session("scode")="admin" then
	'	Response.Write "===========chkFileExist_virtual========end" & "<BR>"
	'end if
end function

'�\���ɮ�
function CoverFile(strpath,local_seq1,local_seq2)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
'	strTestFile = Server.MapPath(strpath)
	objFSO.MoveFile strTestFile&"\"&local_seq2,strTestFile&"\"&local_seq2&"_tmp"
	call renameFile(strpath,local_seq1,local_seq2)
	call DelFile(strpath,local_seq1,local_seq2)
	set objFSO = nothing
end function

'�屼�ɮ�
function DelFile(strpath,local_seq)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
'	strTestFile = Server.MapPath(strpath)
	'Response.Write strpath&local_seq &"<BR>"
	if objFSO.FileExists(strpath&local_seq) = True then
		objFSO.DeleteFile strpath&local_seq
	end if 
	set objFSO = nothing
end function

'�ɮ׭��s�R�W
'function renameFile(strpath,local_seq1,local_seq2)
 '     Dim strTestFile,objFSO
  '    strTestFile = Server.MapPath(strpath)
   '   set objFSO  = CreateObject("Scripting.FileSystemObject")
	 ' if chkFileExist(strpath&local_seq2) = 1 then
	'     set objFile0= objFSO.GetFile(strTestFile&"\"&local_seq2)
	'	 objFile0.Delete
	'  end if
    '  set objFile = objFSO.GetFile(strTestFile&"\"&local_seq1)
'      Response.Write strTestFile&"\"&local_seq2&".gif"
'      Response.end
'      objFile.Move(strTestFile&"\"&local_seq2)
     ' set objFSO=nothing
'end function

'�ɮ׭��s�R�W
function renameFile(strpath,local_seq1,local_seq2)
	Dim strTestFile,objFSO
	strTestFile = Server.MapPath(strpath)
	
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	'if request("ctrlsubmitTask")="C" then
	'else
		if chkFileExist_virtual(strpath&"/"&local_seq2) = 1 then
			set objFile0= objFSO.GetFile(strTestFile&"\"&local_seq2)
			objFile0.Delete
		end if
	'end if
	'Response.Write "strTestFile="& strTestFile &"<BR>"
	'Response.Write "local_seq1="& local_seq1 &"<BR>"
	'Response.Write "local_seq2="& local_seq2 &"<BR>"
	'Response.End 
	set objFile = objFSO.GetFile(strTestFile&"\"&local_seq1)
	objFile.Move(strTestFile&"\"&local_seq2)
	set objFSO=nothing
end function

'�ɮ׭��s�R�W
function renameFile1(strpath1,strpath2)
'strpath1�ӷ���,������|+�ɦW
'strpath2�ت���,������|+�ɦW
	Dim strTestFile1,strTestFile2,objFSO

	'Response.Write "strpath1="& strpath1&"<br>"
	'Response.Write "strpath2="& strpath2&"<br>"
	'response.end
	strTestFile1 = Server.MapPath(strpath1)		
	strTestFile2 = Server.MapPath(strpath2)
	if session("scode")="m983" then
	'	Response.Write strpath1&"<br>"
	'	Response.Write strpath2&"<br>"
	'	Response.Write chkFileExist_virtual(strpath2)&"<br>"
	'	Response.End
	end if
	
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	if request("ctrlsubmitTask")="C" then
	else
		if chkFileExist_virtual(strpath2) = 1 then			'�ˬd�ǵ������|
			set objFile0= objFSO.GetFile(strTestFile2)		'�ǹ�����|
			objFile0.Delete
		end if
	end if

	set objFile = objFSO.GetFile(strTestFile1)			'�ǹ�����|
	objFile.Move(strTestFile2)
	set objFSO=nothing
	exit function
end function
'�ƻs�ɮ�
function copyFile1(strpath1,strpath2)
	'strpath1�ӷ���,������|+�ɦW
	'strpath2�ت���,������|+�ɦW
	Dim strTestFile1,strTestFile2,objFSO
	'response.Write "strpath1="& strpath1 & "<BR>"
	'response.Write "strpath2="& strpath2 & "<BR>"
	'response.End 
	strTestFile1 = Server.MapPath(strpath1)		'�������|�ন������|
	strTestFile2 = Server.MapPath(strpath2)
	'response.Write "strTestFile1="& strTestFile1 & "<BR>"
	'response.Write "strTestFile2="& strTestFile2 & "<BR>"
	'response.End 
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	if chkFileExist_virtual(strpath2) = 1 then			'�ˬd�ǵ������|

	end if
	set objFile = objFSO.GetFile(strTestFile1)			'�ǹ�����|
	objFSO.CopyFile strTestFile1,strTestFile2
	set objFSO=nothing
end function
'�ƻs�ɮ׭��s�R�W
function copyrenameFile(strpath1,strpath2)
	'strpath1�ӷ���,������|+�ɦW
	'strpath2�ت���,������|+�ɦW
	Dim strTestFile1,strTestFile2,objFSO
	'if session("scode")="admin" then
	'	Response.Write "===========copyrenameFile========begin" & "<BR>"
	'end if
	'response.Write "strpath1="& strpath1 & "<BR>"
	'response.Write "strpath2="& strpath2 & "<BR>"
	strTestFile1 = Server.MapPath(strpath1)		
	strTestFile2 = Server.MapPath(strpath2)
	if session("scode")="m983" then
	'    response.Write "strTestFile1="& strTestFile1 & "<BR>"
	'    response.Write "strTestFile2="& strTestFile2 & "<BR>"
	'	Response.Write chkFileExist_virtual(strpath2)&"<br>"
	end if
	
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	if chkFileExist_virtual(strpath2) = 1 then			'�ˬd�ǵ������|
		set objFile0= objFSO.GetFile(strTestFile2)		'�ǹ�����|
		objFile0.Delete
	end if
	'if session("scode")="admin" then
		'Response.Write objFSO.GetFile(strTestFile1) & "<BR>"
		'Response.End
	'end if

	set objFile = objFSO.GetFile(strTestFile1)			'�ǹ�����|
	'if session("scode")="admin" then
	'	Response.Write "strTestFile1="& strTestFile1&"<br>"
	'	Response.Write "strTestFile2="& strTestFile2&"<br>"
	'end if
	objFSO.CopyFile strTestFile1,strTestFile2
	set objFSO=nothing
	'if session("scode")="admin" then
	'	Response.Write "===========copyrenameFile========end" & "<BR>"
	'	Response.End
	'end if

	exit function
end function
'�ƻs�ɮ׭��s�R�W for ���ƻs����
function renameFile2(strpath1,strpath2)
'strpath1�ӷ���,������|+�ɦW
'strpath2�ت���,������|+�ɦW
	Dim strTestFile1,strTestFile2,objFSO

	'Response.Write strpath1&"<br>"
	'Response.Write strpath2&"<br>"
	strTestFile1 = Server.MapPath(strpath1)		
	strTestFile2 = Server.MapPath(strpath2)
	'Response.Write strTestFile1&"<br>"
	'Response.Write strTestFile2&"<br>"
	'Response.Write chkFileExist_virtual(strpath2)&"<br>"
	'Response.End
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	if chkFileExist_virtual(strpath2) = 1 then			'�ˬd�ǵ������|
		'Response.Write "1" &"<BR>"
		set objFile0 = objFSO.GetFile(strTestFile2)		'�ǹ�����|
		objFile0.Delete
	else
		'Response.Write "0" &"<BR>"
	end if
	'if session("scode")="admin" then
	'	Response.Write strTestFile1&"<br>"
	'	Response.Write strTestFile2&"<br>"
	'	Response.End
	'end if

	set objFile = objFSO.GetFile(strTestFile1)			'�ǹ�����|
	objFSO.CopyFile strTestFile1,strTestFile2
	set objFSO=nothing
	exit function
end function
'�N�Y�w�s�b���ɮ׳ƥ�-������|
function copyrenameFile2(strTestFile1,strTestFile2)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTestFile1) then	 
		'�ɮפw�s�b	
		set objFile = objFSO.GetFile(strTestFile1)			'�ǹ�����|
		if trim(strTestFile2)<>empty then
		else
			strTestFile2 = left(strTestFile1,instr(strTestFile1,".")-1)
			strTestFile2 = strTestFile2 & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) &"-"& session("scode") & right (strTestFile1,4)
		end if
		'Response.Write strTestFile1 &"<BR>"
		'Response.Write strTestFile2 &"<BR>"
		'Response.End 
		objFSO.CopyFile strTestFile1,strTestFile2
	end if
	set objFSO = nothing	
end function
'�N�Y�w�s�b���ɮ׳ƥ�-�������|
function copyrenameFile3(strTestFile1,strTestFile2,strTestFile3)
	'Dim strTestFile1,strTestFile2,strTestFile2c
	'response.Write "strTestFile1="& strTestFile1 & "<BR>" '���ɮת����|
	'response.Write "strTestFile2="& strTestFile2 & "<BR>" '�����n�ƻs�ܪ��ɮ׸��|
	'response.Write "strTestFile3="& strTestFile3 & "<BR>" '�ƥ������|
	'response.End 
	if chkFileExist_virtual(strTestFile2) = 1 then			'�ˬd�ǵ������|
	    set objFSO  = CreateObject("Scripting.FileSystemObject")
        strTestFile1 = Server.MapPath(strTestFile1)		'�������|�ন������|
        strTestFile2 = Server.MapPath(strTestFile2)
        strTestFile3 = Server.MapPath(strTestFile3)
	    strTestFile3c = left(strTestFile3,instr(strTestFile3,".")-1)
	    strTestFile3c = strTestFile3c & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) &"-"& session("scode") &"."& mid(strTestFile3,instr(strTestFile3,".")+1)
	    'response.Write "strTestFile2="& strTestFile2 & "<BR>"
        'response.Write "strTestFile3c="& strTestFile3c & "<BR>"
        'response.End 
        set objFile = objFSO.GetFile(strTestFile2)			'�ǹ�����|
        objFSO.CopyFile strTestFile2,strTestFile3c
	    set objFSO=nothing
	end if
end function

%>
