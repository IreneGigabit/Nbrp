<%
'將檔案複製到區所
'pfromfile=來源檔案,ptofile目地檔案
'如pfromfile=/brp/NPE/temp/N/_/166/16678/PE-16678--0000-0-1.jpg
function copyfile_tobranch(pbr_branch,pfromfile)
	
	call getFileServer(pbr_branch)

	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
		'response.write pfromfile &"<BR>"
		'response.write attach_path &"<BR>"
		'response.write session("prgid") &"<BR>"
		if session("prgid") = "brp64" or session("prgid") = "exp64" then '營洽退回承辦
			'來源的路徑
			if left(session("prgid"),3) = "brp" then
				attach_path = replace(pfromfile,"/brp/"&pbr_branch &"P","")
			else
				attach_path = replace(pfromfile,"/brp/"&pbr_branch &"PE","")
			end if
			'attach_path = pfromfile
			'response.write "來源: "& attach_path &"<BR>"
			'目的foldername
			foldername = replace(attach_path,"/temp/"& pbr_branch,"")
			'response.write "目的: " & foldername &"<BR>"
			folder_name = "/temp/"& pbr_branch & left(foldername,instrRev(foldername,"/",-1,1)-1)
			'response.write folder_name &"<BR>"
			'區所的file_name
			filename=mid(attach_path,instrRev(attach_path,"/",-1,1)+1)
			'response.write filename &"<BR>"
		else
			'區所的路徑
			attach_path = replace(pfromfile,"/temp/"&pbr_branch,"")
			'response.write attach_path &"<BR>"
			'區所的foldername
			if left(session("prgid"),3) = "brp" then
				foldername = replace(attach_path,"/brp/"& pbr_branch &"P","")
			else
				foldername = replace(attach_path,"/brp/"& pbr_branch &"PE","")
			end if
			folder_name=left(foldername,instrRev(foldername,"/",-1,1)-1)
			'區所的file_name
			filename=mid(attach_path,instrRev(attach_path,"/",-1,1)+1)
			'response.write filename &"<BR>"
		end if
		'Response.End 
		
		newfilename = gbrWebDir & folder_name &"/"&filename
		newfoldername = gbrWebDir & folder_name &"/"
		'response.write newfilename &"<BR>"
		'response.write newfoldername &"<BR>"
		'response.end

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
		'response.write strTestFile1 &"<BR>"
		'response.write strTestFile2 &"<BR>"
		'response.end
		
		set objFile = objFSO.GetFile(strTestFile1)			'傳實體路徑
		objFSO.CopyFile strTestFile1,strTestFile2
		'File_name_new2 = left(pfromfile,len(pfromfile)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (pfromfile,4)
		'renameFile1 pfromfile,File_name_new2
	set objFSO=nothing
end function
%>
