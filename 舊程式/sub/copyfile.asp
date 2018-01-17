<%
'http://web02/fext/sub/copyfile.asp

'Set MyFile = objFSO.CreateTextFile("c:\testfile.txt", True)
'MyFile.WriteLine("This is a test.")
'MyFile.Close
'strTestFolder = Server.MapPath("web01")

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

'建立目錄
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
		'strTestFolder = pfldr1 & "\" & aryfolder(i)
		'Response.Write "createFolder: "& strTestFolder &"<BR>"
		
		if chkfolderExist(strTestFolder) = 1 then
		else
			objFSO.CreateFolder(strTestFolder)
		end if	
	next 
	'Response.End 
	set objFSO = nothing
end function

'檢查檔案是否存在
function chkFileExist(strFile)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	'strTestFile = Server.MapPath(strFile)
	if objFSO.FileExists(strFile) then
		chkFileExist=1
	else
		chkFileExist=0
	end if
	set objFSO = nothing
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

'將檔案複製到區所
'pfromfile=來源檔案,ptofile目地檔案
'如pfromfile=/brp/NPE/temp/N/_/166/16678/PE-16678--0000-0-1.jpg
function copyfile_tobranch(pbr_branch,pfromfile)
	
	call getFileServer(pbr_branch)

	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
		'response.write pfromfile &"<BR>"

		'區所的路徑
		attach_path=replace(pfromfile,"/temp/"&pbr_branch,"")
		'response.write attach_path &"<BR>"
		'區所的foldername
		foldername=replace(attach_path,"/brp/"& pbr_branch &"PE","")
		folder_name=left(foldername,instrRev(foldername,"/",-1,1)-1)
		
		'區所的file_name
		filename=mid(attach_path,instrRev(attach_path,"/",-1,1)+1)
		'response.write filename &"<BR>"
		'response.end
		
		newfilename = gbrWebDir & folder_name &"/"&filename
		newfoldername = gbrWebDir & folder_name &"/"

	'建立區所所在目錄
		call Check_CreateFolder_virtual(gbrWebDir,folder_name)
	'檢查區所有沒有資料有的話先備份起來起
		if chkFileExist_virtual(newfilename) = 1 then
			'將刪除檔案時是將原來word備份起來(日期時間取另一個名字)
			File_name_new = left(newfilename,len(newfilename)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (newfilename,4)
			renameFile1 newfilename,File_name_new
		end if
	'將資料copy到區所
		'strTestFile1 = Server.MapPath(pfromfile)		
		'strTestFile2 = Server.MapPath(newfilename)
		'set objFile = objFSO.GetFile(pfromfile)			'傳實體路徑
		'objFile.CopyFile strTestFile1,strTestFile2
		strTestFile1 = Server.MapPath(pfromfile)		
		strTestFile2 = Server.MapPath(newfilename)
		
		set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
		objFSO.CopyFile strTestFile1,strTestFile2
		'File_name_new2 = left(pfromfile,len(pfromfile)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (pfromfile,4)
		'renameFile1 pfromfile,File_name_new2
	set objFSO=nothing
end function

'將檔案複製到區所
function copyfile_tobranchxx(pbr_branch,pfile1)
	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
	'brfldr1 = "/fext/NTE_doc/FTE_file"
	'fldr2 = "_/021/021050/"
	'pfile1 = "TE-030005--0001-1.txt"
	'---test
	'brfldr1 = "\\web03\NTE\FTE_File"
	'strTestFolder = "/fext/NTE/FTE_File/_"
	'Response.Write strTestFolder & "<BR>"
	'D:\Data\document\NtE\FTE_File\_\030
	'strTestFolder = server.MapPath(strTestFolder)
	'Response.Write strTestFolder & "<BR>"
	'strTestFolder = strTestFolder & "\030"
	'Response.Write strTestFolder & "<BR>"
	'Response.end
	'objFSO.CreateFolder(strTestFolder)
	'Response.end
	'---end
	
	Select Case Request.ServerVariables("SERVER_NAME")
		Case "web01"
			fileservername = "\\web03"
		Case "web02"
			'fileservername = "\\web02"
			fileservername = "/fext"
		Case else
			fileservername = "\\sin31"	'?上線時需修改
	end Select
	sfldr1 = fileservername & "\FTE_File\"	'國外所來源資料實體路徑
	brfldr1 = fileservername & "\"& pbr_branch & "TE\FTE_File" '區所
	
	if Request.ServerVariables("SERVER_NAME")="web02" then
		sfldr1 = replace(sfldr1,"\","/")
		sfldr1 = server.MapPath(sfldr1)
		brfldr1 = replace(brfldr1,"\","/")
		brfldr1 = server.MapPath(brfldr1)
	end if
	
	'Response.Write "國外所:" & sfldr1&"<br>"
	'Response.Write "區所:" & brfldr1&"<br>"
	arfile1 = split(pfile1,"-")
	if arfile1(2)="" then
		fldr2 = "_/"
	else
		fldr2 = arfile1(2) & "/"
	end if
	fldr2 = fldr2 & left(arfile1(1),3) & "/" & arfile1(1)
	'response.write "fldr2="&fldr2&"<BR>"

	brfldrfall = brfldr1 & "/" & fldr2	'區所完整路徑
	'response.write "區所brfldr1="&brfldr1&"<BR>"
	'response.write "brfldrfall="&brfldrfall&"<BR>"
	'Response.End 
	
	call Check_CreateFolder(brfldr1,fldr2) '建立目錄
	
	if chkFileExist(brfldrfall&pfile1)=1 then '檢查檔案是否存在
		call DelFile(brfldrfall,pfile1)
	end if
	
	sfldrfall = sfldr1 & fldr2 & "\" & pfile1
	'Response.Write fldr2 & "\" & pfile1 &"<BR>"
	if Request.ServerVariables("SERVER_NAME")="web02" then
		sfldrfall = sfldr1 &"/"& fldr2 & "\" & pfile1
		sfldrfall = replace(sfldrfall,"\","/")
		brfldrfall = replace(brfldrfall,"\","/")
	end if
	'Response.Write "***copy file <BR> from: "& sfldrfall &"<BR>"
	'Response.Write "to: " & brfldrfall &"<BR>"
	'Response.End 
	
	objFSO.CopyFile sfldrfall,brfldrfall&"\"

	set objFSO=nothing
end function
%>
