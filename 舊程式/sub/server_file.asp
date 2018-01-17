<%
'定義共用參數
gdept = "P"
gbrfileservername = ""
gbrfilesmapervername = ""  '實際檔案主機
gfileservername = ""
gbrDir = ""
gDir = ""
gbrWebDir = ""
gWebDir = ""
gbrDbDir = ""
gcustDbDir = ""

'取得server name
'取得 (1) gbrFileServerName(區所上傳server name) (2)gFileServerName(國外所上傳Server Name)
'(3) gbrDir(區所檔案的實體路徑ex.\\sinn03\NPE)    (4)gDir(國外所檔案的實體路徑 ex.\\sin31\FPE_File)
'(5) gbrWebDir(區所檔案的虛擬路徑ex./brp/NPE)    (6)gWebDir(國外所檔案的虛擬路徑 ex./fexp/FPE_File)
'(7) gbrDbDir(區所請款單的虛擬路徑ex./brp/brdb_file) (8) gcustDbDir(區所對催帳客函的虛擬路徑ex./brp/custdb_file)
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
				gbrfilesmapervername = "sinn01"  '實際檔案主機
			Case "C" 
				gbrfileservername = "sic09"
				gbrfilesmapervername = "sic08"  '實際檔案主機
			Case "S" 
				gbrfileservername = "sis09"
				gbrfilesmapervername = "sis08"   '實際檔案主機
			Case "K"
				gbrfileservername = "sik09"
				gbrfilesmapervername = "sik08"   '實際檔案主機
			End Select 
			gfileservername = "sin31"	'?上線時需修改
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
	'2012/5/2增加請款單及對催帳客函
	gbrDbDir = "/brp/brdb_file"
	gcustDbDir = "/brp/custdb_file"
	'Response.Write "gbrWebDir=" & gbrWebDir & "<br>"
	'Response.Write "gWebDir=" & gWebDir & "<br>"
	'Response.Write "gbrDir=" & gbrDir & "<br>"
	'Response.Write "gDir=" & gDir & "<br>"
	'Response.end
end function

'檢查磁碟是否存在
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

'檢查目錄是否存在
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

'建立目錄(使用 share folder方式,直接利用實體路徑 create folder)
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

'建立目錄(傳入virtual directory,再用mappath方式轉換成實體路徑,Create folder )
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

'檢查檔案是否存在(傳入的為實體路徑,ex.share folder方式 : //web02/FPE_File)
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

'檢查檔案是否存在(傳入為虛擬目錄 ex./fexp/FPE_File/_/700/70002)
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

'蓋掉檔案
function CoverFile(strpath,local_seq1,local_seq2)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
'	strTestFile = Server.MapPath(strpath)
	objFSO.MoveFile strTestFile&"\"&local_seq2,strTestFile&"\"&local_seq2&"_tmp"
	call renameFile(strpath,local_seq1,local_seq2)
	call DelFile(strpath,local_seq1,local_seq2)
	set objFSO = nothing
end function

'砍掉檔案
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

'檔案重新命名
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

'檔案重新命名
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

'檔案重新命名
function renameFile1(strpath1,strpath2)
'strpath1來源檔,完整路徑+檔名
'strpath2目的檔,完整路徑+檔名
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
		if chkFileExist_virtual(strpath2) = 1 then			'檢查傳虛擬路徑
			set objFile0= objFSO.GetFile(strTestFile2)		'傳實體路徑
			objFile0.Delete
		end if
	end if

	set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
	objFile.Move(strTestFile2)
	set objFSO=nothing
	exit function
end function
'複製檔案
function copyFile1(strpath1,strpath2)
	'strpath1來源檔,完整路徑+檔名
	'strpath2目的檔,完整路徑+檔名
	Dim strTestFile1,strTestFile2,objFSO
	'response.Write "strpath1="& strpath1 & "<BR>"
	'response.Write "strpath2="& strpath2 & "<BR>"
	'response.End 
	strTestFile1 = Server.MapPath(strpath1)		'虛擬路徑轉成實體路徑
	strTestFile2 = Server.MapPath(strpath2)
	'response.Write "strTestFile1="& strTestFile1 & "<BR>"
	'response.Write "strTestFile2="& strTestFile2 & "<BR>"
	'response.End 
	set objFSO  = CreateObject("Scripting.FileSystemObject")
	if chkFileExist_virtual(strpath2) = 1 then			'檢查傳虛擬路徑

	end if
	set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
	objFSO.CopyFile strTestFile1,strTestFile2
	set objFSO=nothing
end function
'複製檔案重新命名
function copyrenameFile(strpath1,strpath2)
	'strpath1來源檔,完整路徑+檔名
	'strpath2目的檔,完整路徑+檔名
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
	if chkFileExist_virtual(strpath2) = 1 then			'檢查傳虛擬路徑
		set objFile0= objFSO.GetFile(strTestFile2)		'傳實體路徑
		objFile0.Delete
	end if
	'if session("scode")="admin" then
		'Response.Write objFSO.GetFile(strTestFile1) & "<BR>"
		'Response.End
	'end if

	set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
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
'複製檔案重新命名 for 交辦複製附件
function renameFile2(strpath1,strpath2)
'strpath1來源檔,完整路徑+檔名
'strpath2目的檔,完整路徑+檔名
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
	if chkFileExist_virtual(strpath2) = 1 then			'檢查傳虛擬路徑
		'Response.Write "1" &"<BR>"
		set objFile0 = objFSO.GetFile(strTestFile2)		'傳實體路徑
		objFile0.Delete
	else
		'Response.Write "0" &"<BR>"
	end if
	'if session("scode")="admin" then
	'	Response.Write strTestFile1&"<br>"
	'	Response.Write strTestFile2&"<br>"
	'	Response.End
	'end if

	set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
	objFSO.CopyFile strTestFile1,strTestFile2
	set objFSO=nothing
	exit function
end function
'將若已存在的檔案備份-實體路徑
function copyrenameFile2(strTestFile1,strTestFile2)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTestFile1) then	 
		'檔案已存在	
		set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
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
'將若已存在的檔案備份-虛擬路徑
function copyrenameFile3(strTestFile1,strTestFile2,strTestFile3)
	'Dim strTestFile1,strTestFile2,strTestFile2c
	'response.Write "strTestFile1="& strTestFile1 & "<BR>" '原檔案的路徑
	'response.Write "strTestFile2="& strTestFile2 & "<BR>" '本次要複製至的檔案路徑
	'response.Write "strTestFile3="& strTestFile3 & "<BR>" '備份的路徑
	'response.End 
	if chkFileExist_virtual(strTestFile2) = 1 then			'檢查傳虛擬路徑
	    set objFSO  = CreateObject("Scripting.FileSystemObject")
        strTestFile1 = Server.MapPath(strTestFile1)		'虛擬路徑轉成實體路徑
        strTestFile2 = Server.MapPath(strTestFile2)
        strTestFile3 = Server.MapPath(strTestFile3)
	    strTestFile3c = left(strTestFile3,instr(strTestFile3,".")-1)
	    strTestFile3c = strTestFile3c & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) &"-"& session("scode") &"."& mid(strTestFile3,instr(strTestFile3,".")+1)
	    'response.Write "strTestFile2="& strTestFile2 & "<BR>"
        'response.Write "strTestFile3c="& strTestFile3c & "<BR>"
        'response.End 
        set objFile = objFSO.GetFile(strTestFile2)			'傳實體路徑
        objFSO.CopyFile strTestFile2,strTestFile3c
	    set objFSO=nothing
	end if
end function

%>
