<%
'�N�ɮ׽ƻs��ϩ�
'pfromfile=�ӷ��ɮ�,ptofile�ئa�ɮ�
'�ppfromfile=/brp/NPE/temp/N/_/166/16678/PE-16678--0000-0-1.jpg
function copyfile_tobranch(pbr_branch,pfromfile)
	
	call getFileServer(pbr_branch)

	Dim strTestFolder,objFSO
	set objFSO = CreateObject("Scripting.FileSystemObject")
		'response.write pfromfile &"<BR>"
		'response.write attach_path &"<BR>"
		'response.write session("prgid") &"<BR>"
		if session("prgid") = "brp64" or session("prgid") = "exp64" then '�笢�h�^�ӿ�
			'�ӷ������|
			if left(session("prgid"),3) = "brp" then
				attach_path = replace(pfromfile,"/brp/"&pbr_branch &"P","")
			else
				attach_path = replace(pfromfile,"/brp/"&pbr_branch &"PE","")
			end if
			'attach_path = pfromfile
			'response.write "�ӷ�: "& attach_path &"<BR>"
			'�ت�foldername
			foldername = replace(attach_path,"/temp/"& pbr_branch,"")
			'response.write "�ت�: " & foldername &"<BR>"
			folder_name = "/temp/"& pbr_branch & left(foldername,instrRev(foldername,"/",-1,1)-1)
			'response.write folder_name &"<BR>"
			'�ϩҪ�file_name
			filename=mid(attach_path,instrRev(attach_path,"/",-1,1)+1)
			'response.write filename &"<BR>"
		else
			'�ϩҪ����|
			attach_path = replace(pfromfile,"/temp/"&pbr_branch,"")
			'response.write attach_path &"<BR>"
			'�ϩҪ�foldername
			if left(session("prgid"),3) = "brp" then
				foldername = replace(attach_path,"/brp/"& pbr_branch &"P","")
			else
				foldername = replace(attach_path,"/brp/"& pbr_branch &"PE","")
			end if
			folder_name=left(foldername,instrRev(foldername,"/",-1,1)-1)
			'�ϩҪ�file_name
			filename=mid(attach_path,instrRev(attach_path,"/",-1,1)+1)
			'response.write filename &"<BR>"
		end if
		'Response.End 
		
		newfilename = gbrWebDir & folder_name &"/"&filename
		newfoldername = gbrWebDir & folder_name &"/"
		'response.write newfilename &"<BR>"
		'response.write newfoldername &"<BR>"
		'response.end

		'�إ߰ϩҩҦb�ؿ�
		call Check_CreateFolder_virtual(gbrWebDir,folder_name)
		'�ˬd�ϩҦ��S����Ʀ����ܥ��ƥ��_�Ӱ_
		if chkFileExist_virtual(newfilename) = 1 then
			'�N�R���ɮ׮ɬO�N���word�ƥ��_��(����ɶ����t�@�ӦW�r)
			File_name_new = left(newfilename,len(newfilename)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (newfilename,4)
			renameFile1 newfilename,File_name_new
		end if
		'�N���copy��ϩ�
		'strTestFile1 = Server.MapPath(pfromfile)		
		'strTestFile2 = Server.MapPath(newfilename)
		'set objFile = objFSO.GetFile(pfromfile)			'�ǹ�����|
		'objFile.CopyFile strTestFile1,strTestFile2
		strTestFile1 = Server.MapPath(pfromfile)		
		strTestFile2 = Server.MapPath(newfilename)
		'response.write strTestFile1 &"<BR>"
		'response.write strTestFile2 &"<BR>"
		'response.end
		
		set objFile = objFSO.GetFile(strTestFile1)			'�ǹ�����|
		objFSO.CopyFile strTestFile1,strTestFile2
		'File_name_new2 = left(pfromfile,len(pfromfile)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (pfromfile,4)
		'renameFile1 pfromfile,File_name_new2
	set objFSO=nothing
end function
%>
