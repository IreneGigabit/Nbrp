<% Server.ScriptTimeout = 2400 %>

<html>
<head>
<title><%=cont%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<!--#include file="class.upload.asp"--><!--自訂物件檔-->
<!--#INCLUDE FILE="../sub/Server_File.asp" -->
<%if session("filename_flag")="pic" then%>
<!--#Include file="../sub/Server_savelog.vbs" -->
<%end if%>
<%
response.buffer=true 
'mappath_name="\Fexp\FPE_file\"

'取得 (1) gbrFileServerName(區所上傳server name) (2)gFileServerName(國外所上傳Server Name)
'(3) gbrDir(區所檔案的實體路徑ex.\\sin09\NPE)    (4)gDir(國外所檔案的實體路徑 ex.\\sin31\FPE_File)
'(3) gbrWebDir(區所檔案的虛擬路徑ex./brp/NPE)    (4)gWebDir(國外所檔案的虛擬路徑 ex./fexp/FPE_File)
'(7) gbrDbDir(區所請款單的虛擬路徑ex./brp/brdb_file) (8) gcustDbDir(區所對催帳客函的虛擬路徑ex./brp/custdb_file)
'Response.Write session("docbranch") & "<BR>"
IF session("docbranch")<>empty then
	call getFileServer(session("docbranch"))
	Session("folder_name")="temp/"& session("docbranch") & "/" & session("folder_name")
Else
	call getFileServer(session("se_branch"))
End IF
'Response.Write session("folder_name") & "<BR>"
'Response.End

Dim mySmartUpload
select case lcase(session("type"))
	case "doc"
		seq = session("seq")
		seq1 = session("seq1")
		folder_name = session("folder_name")
		'if session("scode")="m983" then Response.Write "session--folder_name="& folder_name & "<BR>"
		prefix_name = session("prefix_name")
		draw_file = session("draw_file")
		old_file = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
		if session("filename_flag")="pic" then
			file_path = session("picturepath"&session("se_branch")) & "/" & folder_name
		else
			'file_path="\FEXP\FPE_file\" & folder_name
			file_path = gbrWebDir & "/" & folder_name
		end if
		'if session("scode")="m983" then Response.Write "gbrWebDir="& gbrWebDir & "<BR>"
		'if session("scode")="m983" then Response.Write "file_path="& file_path & "<BR>"
		'Response.End
	   
		call Check_CreateFolder_virtual(gbrWebDir,folder_name)
	    '回傳檔案的欄位名
		form_name="reg."&session("form_name")
		gsize_name = "reg."&session("size_name")
		gfile_name = "reg."&session("file_name")
		gsource_name = "reg."&session("source_name")
		gbtnname = "reg."&session("btnname")
		doc_in_scode = "reg."&session("doc_in_scode")
		doc_in_date = "reg."&session("doc_in_date")
		'Response.write gpath_name
		'Response.end
	case "photo"
		seq = session("seq")
		draw_file = session("draw_file")
		old_file = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
		if instr(session("seq"),"-") > 0 then
			ar = split(session("seq"),"-")
			lForder1 = mid(right("0000"&ar(0),5),1,1)
			lForder2 = mid(right("0000"&ar(0),5),2,2)
		else
			lForder1 = mid(right("0000"&session("seq"),5),1,1)
			lForder2 = mid(right("0000"&session("seq"),5),2,2)
		end if	
		file_path="/FTE_file/"&lForder1&"/"&lForder2
		'檢查Folder是否存在
		if chkfolderExist("/FTE_file/"&lForder1) = 1 then
		else
			call CreateFolder("/FTE_file/"&lForder1)
		end if
		if chkfolderExist("/FTE_file/"&lForder1&"/"&lForder2) = 1 then
		else
			call CreateFolder("/FTE_file/"&lForder1&"/"&lForder2)
		end if
	case "apcust_file","custdb_file","db_file","custresp_file","brdb_file"
		'2012/5/2增加，custdb_file=對催帳客函 db_file=請款單 custresp_file=客戶對催回應文件
		'2015/11/16 apcust_file 契約書、委任書
		'2016/11/22增加brdb_file=英文invoice
	   	folder_name = session("folder_name")
		prefix_name = session("prefix_name")
		draw_file = session("draw_file")
		old_file = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
	   'file_path="\FEXP\FPE_file\" & folder_name
	   if lcase(session("type"))="apcust_file" or lcase(session("type"))="custdb_file" or lcase(session("type"))="custresp_file" then
	      file_path = gcustDbDir & "/" & folder_name
	      'Response.Write "gcustDbDir="& gcustDbDir & "<br>"
	      'Response.Write "folder_name="& folder_name &"<br>"
	      'Response.End
	      call Check_CreateFolder_virtual(gcustDbDir,folder_name)
	   else
	      file_path = gbrDbDir & "/" & folder_name
	      call Check_CreateFolder_virtual(gbrDbDir,folder_name)
	   end if   
	   
	   'Response.Write "gDebitDir="& gDebitDir & "<br>"
	   'Response.Write "folder_name="& folder_name &"<br>"
       'Response.End
       
	    '回傳檔案的欄位名
		form_name="reg."&session("form_name")
		gsize_name = "reg."&session("size_name")
		gfile_name = "reg."&session("file_name")
		gsource_name = "reg."&session("source_name")
		gbtnname = "reg."&session("btnname")
		doc_in_scode = "reg."&session("doc_in_scode")
		doc_in_scodenm = "reg."&session("doc_in_scodenm")
		doc_in_date = "reg."&session("doc_in_date")
		gdraw_name = "reg."&session("draw_name")
		if lcase(session("type"))="apcust_file" then
		else
		    db_file_flag = "reg."&session("db_file_flag")
		end if
		prgid_name = "reg."&session("prgid_name")
		attach_flag_name = "reg." & session("attach_flag_name")	
		'Response.write "db="& db_file_flag
		'response.Write "form_name="& form_name & "<BR>"
		'Response.end	
	case else
		Response.redirect "upload_win.asp"
end select

Server.ScriptTimeOut = 1200

'重新上傳
Set mySmartUpload = Server.CreateObject("UpDownExpress.FileUpload")
Set up = New Upload      '建立Upload物件
Dim strTestFile,objFSO

'Response.Write file_path &"<BR>"
'up.path = replace(file_path,"/","\")      '指定儲存路徑
file_path = replace(file_path,"\","/") 
'Response.Write "file_path= " & file_path & "<br>"

up.path = file_path      '指定儲存路徑
dd=right(up.get_file1("theFile"),len(up.get_file1("theFile"))-InstrRev(up.get_file1("theFile"),"\"))
'Response.Write "dd= " & dd & "<br>"

'Response.Write "filename_flag= " & session("filename_flag") & "<br>"
if session("filename_flag")="source_name" then
	ee = up.get_file(up.get_path("theFile")) 
	'原始檔名
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
	'原始檔名
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
	'原始檔名
	original_name=up.get_file(up.get_path("theFile"))
	ee = session("nfilename") &"_"& original_name
else
    '新檔名
	ee = session("nfilename") &"."& right(up.get_file1("theFile"),len(up.get_file1("theFile"))-InstrRev(up.get_file1("theFile"),"."))
	'原始檔名
	original_name=up.get_file(up.get_path("theFile"))
end if

bb = up.get_file(up.get_path("theFile"))
'old_file1=up.get_file(up.get_path("theFile"))
old_file1 = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))

if session("scode")="m983" then
    Response.Write "file_path= " & file_path & "<br>"
    Response.Write "ee="& ee &"<BR>"
end if
'Response.Write "nfilename="& session("nfilename") &"<BR>"
'Response.Write "original_name="& original_name &"<BR>"
'Response.Write "bb="& bb &"<BR>"
'Response.Write "old_file1="& old_file1 &"<BR>"
'Response.Write "原始檔名 ee=" & ee & "<br>"
'Response.Write "dd="& dd &"<BR>"
'Response.Write "filename_flag="& session("filename_flag") &"<BR>"
'if session("type")="custresp_file" then
'Response.End 
'end if				  
if bb = old_file1 then
	'Response.Write "aaaa"
	'Response.end
	'2012/5/2增加，若請款單或對催帳客函之檔案已存在，則先備份再覆蓋
	'2015/11/16 apcust_file 契約書、委任書
	if session("type")="apcust_file" or session("type")="custdb_file" or session("type")="db_file" or session("type")="custresp_file" or session("type")="brdb_file" then
	   File_name=file_path&"/"&old_file1 
	   '備份名字規則：檔名_年月日時分秒
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
    'Response.Write "ee="& ee &"<BR>"
    'Response.End 
	'檢查上傳的檔案是否存在
	Response.Write "<SCRIPT LANGUAGE=vbs>"& chr(13)
	Response.Write " alert(""此檔案已存在! 已覆蓋檔案."")"& chr(13)
	Response.Write "</SCRIPT>"& chr(13)
else
	'Response.Write "bbb"
	'Response.end
	'2012/5/2增加，因請款單及對催帳客函傳入路徑為需擬路徑/brp/custdb_file，所以另外判斷
    if session("type")="custdb_file" or session("type")="db_file" or session("type")="custresp_file" then
       if chkFileExist_virtual(file_path&"/"&dd)=1 then
          attach_flag_value="U"
          ee=""%>
		    <SCRIPT LANGUAGE=vbs>
			msgbox "該檔案已經存在!!" & chr(10) & chr(10) & "請將該檔案更名，並重新上傳!!" 
			window.close
			</SCRIPT>
	  <%else
		    attach_flag_value="A"
	        a = up.SaveFile("theFile") 
			tfilename = bb
			tsize = up.get_FileSize("theFile") 
		    aa = file_path&"/"&ee
			if session("filename_flag")="newname" then '使用新檔名
			else
			    'if session("filename_flag")="source_name" then
				    ee = original_name
			    'else
			    '	if ee <> bb then
			    '		call renameFile(file_path,up.get_file(up.get_path("theFile")),ee)
			    '	end if
			    'end if
			end if
	  	end if	
        'Response.Write "aa="& aa &"<BR>"
        'Response.Write "ee="& ee &"<BR>"
        'Response.End 
    else
		if session("filename_flag")="pic" then
			chkfilename = file_path&"/"&ee
		else
			chkfilename = file_path&"/"&dd
		end if
		'Response.Write "chkfilename=" & chkfilename &"<BR>"
		if chkFileExist(chkfilename) = 1 then%>
			<SCRIPT LANGUAGE=vbs>
			msgbox "該檔案已經存在!!" & chr(10) & chr(10) & "請將該檔案更名，並重新上傳!!" 
			window.close
			</SCRIPT>
		<%Else
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
		
		if session("filename_flag")="pic" then
			'改代表圖
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.Open session("btbrtdb")
			conn.BeginTrans 
			
			call insert_log_table(conn,"U",prgid,session("tablename"),"seq;seq1",session("seq")&";"&session("seq1"))
			usql = "update "& session("tablename") &" set "
			usql = usql & "pic_file_branch='"& session("se_branch") &"',pic_file1_path='"& ee &"'"
			usql = usql & " where seq="& session("seq") &" and seq1='"& session("seq1") &"'"
			'Response.Write usql &"<Br>"
			conn.Execute usql
			
			usql = "insert into pic_dmp_log(branch,br_branch,pic_sqlno,spbranch,seq,seq1,num_key,cust_area,cust_seq,cust_name,scode1,pr_scode,pr_date"
			usql = usql & ",case_no,pic_arcase,case0,pic_scode,pic_branch,pic_capname,cust_prod,custprod_no,last_date,expect_num,draft_date"
			usql = usql & ",hope_date,pic_status,pic_remark,alw_date,pic_content,div_scode,div_date,num_remark,hr_cad,hr_div,hr_process"
			usql = usql & ",hr_repair,hr_copy,hr_stay,hr_likeness,hr_type,hr_other,hr_sd,hr_date,pic_num,beg_date,sta_date,pic_record"
			usql = usql & ",phr_status,hr_remark,mark,tran_date,tran_scode,in_date,country,pic_type,del_remark,pic_date,early_date"
			usql = usql & ",ehr_status,dept,filepath,back_remark,edit_date,pre_sqlno,br_sqlno,br_back_remark,file1,file2,file_remark"
			usql = usql & ",tseq,tseq1,arcase,brpic_branch,enger_scode,pr_remark,pr_seq,pr_seq1,rs_class,case1,from_flag,ud_flag,ud_date,ud_scode)"
			usql = usql & " select branch,br_branch,pic_sqlno,spbranch,seq,seq1,num_key,cust_area,cust_seq,cust_name,scode1,pr_scode,pr_date"
			usql = usql & ",case_no,pic_arcase,case0,pic_scode,pic_branch,pic_capname,cust_prod,custprod_no,last_date,expect_num,draft_date"
			usql = usql & ",hope_date,pic_status,pic_remark,alw_date,pic_content,div_scode,div_date,num_remark,hr_cad,hr_div,hr_process"
			usql = usql & ",hr_repair,hr_copy,hr_stay,hr_likeness,hr_type,hr_other,hr_sd,hr_date,pic_num,beg_date,sta_date,pic_record"
			usql = usql & ",phr_status,hr_remark,mark,tran_date,tran_scode,in_date,country,pic_type,del_remark,pic_date,early_date"
			usql = usql & ",ehr_status,dept,filepath,back_remark,edit_date,pre_sqlno,br_sqlno,br_back_remark,file1,file2,file_remark"
			usql = usql & ",tseq,tseq1,arcase,brpic_branch,enger_scode,pr_remark,pr_seq,pr_seq1,rs_class,case1,from_flag"
			usql = usql & ",'U',getdate(),'"& session("scode") &"'"
			usql = usql & " from pic_dmp"
			usql = usql & " where seq=" & session("seq") &" and seq1='"& session("seq1") &"' and pic_sqlno='"& session("pic_sqlno") &"'"
			Response.Write usql &"<Br>"
			conn.Execute usql
			
			usql = "update pic_dmp set "
			usql = usql & "pic_branch='"& session("se_branch") &"',file1='"& ee &"'"
			usql = usql & " where seq=" & session("seq") &" and seq1='"& session("seq1") &"' and pic_sqlno='"& session("pic_sqlno") &"'"
			Response.Write usql &"<Br>"
			conn.Execute usql
			
			if err.number = 0 and conn.errors.count=0 then
				conn.CommitTrans 
			else
				conn.RollbackTrans 
			end if
		end if	
		if session("type")="brdb_file" then '英文invoice,檔名命名規則：E+branch+dept+ar_no副檔名為使用者上傳
			'備份名字規則：檔名_年月日時分秒
			if chkFileExist_virtual(file_path&"/"&old_file1)=1 then
				File_name_new = left(old_file1,len(old_file1)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (old_file1,4)
				call renameFile(file_path,old_file1,File_name_new)
			end if
			
			newfile="E" & session("se_branch") & session("qs_dept") & session("ar_no") & "." & Right(bb,Instr(StrReverse(bb),".")-1)
			attach_flag_value="A"
			a=up.SaveFile("theFile") 
			tfilename=bb
			tsize=up.get_FileSize("theFile")
			aa=file_path&"/"&newfile
			call renameFile(file_path,up.get_file(up.get_path("theFile")),newfile)
			ee = newfile
		end if
	end if	
End IF	
%>
<script language="vbscript">
'msgbox "<%=form_name%>"
'msgbox "aa="&"<%=aa%>"
'msgbox "ee="&"<%=ee%>"
    <%if session("filename_flag")="pic" then%>
		window.opener.resetForm
    <%else%>
	    window.opener.<%=form_name%>.value="<%=aa%>"  '上傳新路徑
		if "<%=gsize_name%>" <> empty then
			window.opener.<%=gsize_name%>.value="<%=tsize%>"
		end if
		if "<%=gFile_Name%>" <> empty then	
			window.opener.<%=gfile_name%>.value="<%=ee%>"  '新檔案
		end if
		if "<%=gbtnname%>" <> empty then	
			window.opener.<%=gbtnname%>.disabled = true
		end if
		IF "<%=doc_in_date%>" <> empty then
			window.opener.<%=doc_in_date%>.value="<%=date()%>"
		End IF
		IF "<%=doc_in_scode%>" <> empty then
			window.opener.<%=doc_in_scode%>.value = "<%=session("scode")%>"
		End IF
		'原始檔案名稱
		if "<%=gsource_Name%>" <> empty then	
			window.opener.<%=gsource_Name%>.value = "<%=original_name%>"
		end if
	<%end if%>
	<%if gdraw_name <> empty and gdraw_name<> "" and gdraw_name<> "reg." then	%>
		window.opener.<%=gdraw_name%>.value = "<%=newfile%>"
	<%end if%>
	<%If db_file_flag <> empty  then %>	
		'2012/5/2 將對催帳客函或請款單產生方式改為「使用者自行上傳」
		for i=0 to window.opener.<%=db_file_flag%>.length-1
			If window.opener.<%=db_file_flag%>(i).value = "Y" then
				execute "window.opener.<%=db_file_flag%>(i).checked = true"
			End If
		next
		<%If gbtnname<>Empty then%>	
		    <%if session("type")="custresp_file" then %>
		    window.opener.<%=gbtnname%>.disabled = true
		    <%else%>
			window.opener.<%=gbtnname%>.disabled = false
			<%end if%>
		<%End If%>	
		if "<%=attach_flag_name%>" <> empty then
		    <%if attach_flag_value<>empty then%>
				window.opener.<%=attach_flag_name%>.value="<%=attach_flag_value%>"
		    <%else%>
			    window.opener.<%=attach_flag_name%>.value="A"
			<%end if%>    
		end if  	
		IF "<%=doc_in_scodenm%>" <> empty then
			window.opener.<%=doc_in_scodenm%>.value = "<%=session("sc_name")%>"
		End IF
	<%End If	%>
	<%if session("scode")="m1583" then
		'response.end
	end if%>
    window.close()   
</script>
</head>
<body bgcolor=#ffffff>
</body>
</html>