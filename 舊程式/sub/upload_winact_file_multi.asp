<%@ Language=VBScript %>
<% Server.ScriptTimeout = 2400 %>
<!--#include file="class.upload.asp"-->
<!--#INCLUDE FILE="../sub/Server_File.asp" -->
<%
Response.ContentType = "text/plain"
'response.Write Request.Form("seq") &"<Br>"
'response.Write Request.Form("seq1") &"<Br>"
'response.End 

'mappath_name="\Fexp\FPE_file\"

'���o (1) gbrFileServerName(�ϩҤW��server name) (2)gFileServerName(��~�ҤW��Server Name)
'(3) gbrDir(�ϩ��ɮת�������|ex.\\sinn01\NPE)    (4)gDir(��~���ɮת�������| ex.\\sin31\FPE_File)
'(3) gbrWebDir(�ϩ��ɮת��������|ex./brp/NPE)    (4)gWebDir(��~���ɮת��������| ex./fexp/FPE_File)
'(7) gbrDbDir(�ϩҽдڳ檺�������|ex./brp/brdb_file) (8) gcustDbDir(�ϩҹ�ʱb�Ȩ窺�������|ex./brp/custdb_file)
'Response.Write session("docbranch") & "<BR>"
call getFileServer(session("se_branch"))
'Response.Write session("folder_name") & "<BR>"
'Response.End

Dim mySmartUpload
prgid = Request("prgid")
uploadfield = request("uploadfield")
seq = request("seq")
seq1 = request("seq1")
step_grade = request("step_grade")
job_sqlno = request("job_sqlno")

tseq = string(5-len(seq),"0") & seq
folder_name = seq1 &"/"& left(tseq,3) &"/"& tseq

'Response.Write "prgid="& prgid & "<BR>"
'Response.Write "seq="& seq & "<BR>"
'Response.Write "seq1="& seq1 & "<BR>"
'Response.Write "step_grade="& step_grade & "<BR>"
'Response.Write "job_sqlno="& job_sqlno & "<BR>"

Response.Write "1#@#aa#@#" + prgid + "#@#" + seq + "#@#" + seq1 + "#@#" + step_grade + "#@#" + job_sqlno + "#@##@#" + sLink

Response.End



seq = session("seq")
seq1 = session("seq1")
folder_name = session("folder_name")  '_/123/12345
prefix_name = session("prefix_name")
draw_file = session("draw_file")  '�ɮצW�� NP-12345--0001-24306-1.doc
old_file = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
file_path = gbrWebDir & "/" & folder_name   'file_path="\FEXP\FPE_file\" & folder_name
'Response.Write "prgid="& session("prgid") & "<BR>"
'Response.Write gbrWebDir & "<BR>"
'Response.Write file_path & "<BR>"
'Response.End

call Check_CreateFolder_virtual(gbrWebDir,folder_name)
'�^���ɮת����W
form_name="reg."&session("form_name")
gsize_name = "reg."&session("size_name")
gfile_name = "reg."&session("file_name")
gsource_name = "reg."&session("source_name")
gbtnname = "reg."&session("btnname")
doc_in_scode = "reg."&session("doc_in_scode")
doc_in_date = "reg."&session("doc_in_date")
'Response.write gpath_name
'Response.end

'���s�W��
Set mySmartUpload = Server.CreateObject("UpDownExpress.FileUpload")
Set up = New Upload      '�إ�Upload����
Dim strTestFile,objFSO

'Response.Write file_path &"<BR>"
'up.path = replace(file_path,"/","\")      '���w�x�s���|
file_path = replace(file_path,"\","/") 

up.path = file_path      '���w�x�s���|
dd=right(up.get_file1("theFile"),len(up.get_file1("theFile"))-InstrRev(up.get_file1("theFile"),"\"))

'Response.Write "filename_flag= " & session("filename_flag") & "<br>"
if session("filename_flag")="source_name" then
	ee = up.get_file(up.get_path("theFile")) 
	'��l�ɦW
	original_name=up.get_file(up.get_path("theFile"))
elseif session("filename_flag")="source_name2" then
	'if session("scode")="admin" then
		'Response.Write "nfilename="& session("nfilename") &"<BR>"
		ee = ""
		if session("nfilename")<>empty then
			ee = session("nfilename") & "-"
		end if
		ee = ee & up.get_file(up.get_path("theFile")) 
	'else
	'	ee = session("nfilename") & "-"
	'	ee = ee & up.get_file(up.get_path("theFile")) 
	'end if
	'��l�ɦW
	'Response.Write "prgid="& session("prgid") &"<BR>"
	if left(session("prgid"),3) = "brp" or left(session("prgid"),3) = "brp" then
		original_name = up.get_file(up.get_path("theFile"))
	else
		if session("nfilename")<>empty then
			original_name = session("nfilename") & "-" & up.get_file(up.get_path("theFile"))
		else
			original_name = up.get_file(up.get_path("theFile"))
		end if
	end if
elseif session("filename_flag")="pic" then
	'ee = up.get_file(up.get_path("theFile")) 
	'��l�ɦW
	original_name=up.get_file(up.get_path("theFile"))
	ee = session("nfilename") &"_"& original_name
else
	ee = session("nfilename") &"."& right(up.get_file1("theFile"),len(up.get_file1("theFile"))-InstrRev(up.get_file1("theFile"),"."))
	'��l�ɦW
	original_name=up.get_file(up.get_path("theFile"))
end if

bb = up.get_file(up.get_path("theFile"))
'old_file1=up.get_file(up.get_path("theFile"))
old_file1 = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))

'Response.Write "file_path= " & file_path & "<br>"
'Response.Write "ee="& ee &"<BR>"
'Response.Write "nfilename="& session("nfilename") &"<BR>"
'Response.Write "original_name="& original_name &"<BR>"
'Response.Write "bb="& bb &"<BR>"
'Response.Write "old_file1="& old_file1 &"<BR>"
'Response.Write "��l�ɦW ee=" & ee & "<br>"
'Response.Write "dd="& dd &"<BR>"
'Response.Write "filename_flag="& session("filename_flag") &"<BR>"
'if session("type")="custresp_file" then
'Response.End 
'end if				  
if bb = old_file1 then
	'Response.Write "aaaa"
	'Response.end
	'2012/5/2�W�[�A�Y�дڳ�ι�ʱb�Ȩ礧�ɮפw�s�b�A�h���ƥ��A�л\
	if session("type")="custdb_file" or session("type")="db_file" or session("type")="custresp_file" then
	   File_name=file_path&"/"&old_file1 
	   '�ƥ��W�r�W�h�G�ɦW_�~���ɤ���
	   File_name_new = left(old_file1,len(old_file1)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (old_file1,4)
	   renameFile file_path,old_file1,File_name_new
	   attach_flag_value="AR"
	end if
	a=up.SaveFile("theFile") 
	tfilename=bb
	tsize=up.get_FileSize("theFile")
	'aa=file_path&"\"&bb
	aa=file_path&"/"&bb
	if session("filename_flag")="source_name" then
		ee = original_name
	else
		call renameFile(file_path,up.get_file1("theFile"),ee)
	end if
	'�ˬd�W�Ǫ��ɮ׬O�_�s�b
	Response.Write "<SCRIPT LANGUAGE=vbs>"& chr(13)
	Response.Write " alert(""���ɮפw�s�b! �w�л\�ɮ�."")"& chr(13)
	Response.Write "</SCRIPT>"& chr(13)
else
	'Response.Write "bbb"
	'Response.end
	'2012/5/2�W�[�A�]�дڳ�ι�ʱb�Ȩ�ǤJ���|���������|/brp/custdb_file�A�ҥH�t�~�P�_
    if session("type")="custdb_file" or session("type")="db_file" or session("type")="custresp_file" then
       if chkFileExist_virtual(file_path&"/"&dd)=1 then
          attach_flag_value="U"
          ee=""%>
		    <SCRIPT LANGUAGE=vbs>
			msgbox "���ɮפw�g�s�b!!" & chr(10) & chr(10) & "�бN���ɮק�W�A�í��s�W��!!" 
			window.close
			</SCRIPT>
	  <%else
		    attach_flag_value="A"
	        a=up.SaveFile("theFile") 
			tfilename=bb
			tsize=up.get_FileSize("theFile")
			aa=file_path&"/"&ee
			'if session("filename_flag")="source_name" then
				ee = original_name
			'else
			'	if ee <> bb then
			'		call renameFile(file_path,up.get_file(up.get_path("theFile")),ee)
			'	end if
			'end if
	  	end if	
    else
		if session("filename_flag")="pic" then
			chkfilename = file_path&"/"&ee
		else
			chkfilename = file_path&"/"&dd
		end if
		'Response.Write "chkfilename=" & chkfilename &"<BR>"
		if chkFileExist(chkfilename) = 1 then
		    '"���ɮפw�g�s�b!! �бN���ɮק�W�A�í��s�W��!!" 
		Else
			'Response.Write "cccc"
			'Response.end

			a=up.SaveFile("theFile") 
			tfilename=bb
			tsize=up.get_FileSize("theFile")
			aa=file_path&"/"&ee
			if session("scode")="admin" then
			'	Response.Write "new file name=" & ee &"<BR>"
			'	Response.Write "source file name=" & up.get_file(up.get_path("theFile")) &"<BR>"
			'	Response.Write file_path &"<BR>"
			'	Response.Write "filename_flag="& session("filename_flag") &"<BR>"
			'	Response.end
			end if
		
			if session("filename_flag")="source_name" then
				ee = original_name
			else
				if ee <> bb then
					call renameFile(file_path,up.get_file(up.get_path("theFile")),ee)
				end if
			end if
		End IF	
		
	end if	
End IF	
%>
