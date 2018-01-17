<!--#INCLUDE FILE="../sub/Server_File.asp" -->
<%if session("filename_flag")="pic" then%>
<!--#Include file="../sub/Server_savelog.vbs" -->
<%end if%>
<%
response.buffer=true 
'取得 (1) gbrFileServerName(區所上傳server name) (2)gFileServerName(國外所上傳Server Name)
'(3) gbrDir(區所檔案的實體路徑ex.\\sinn01\NPE)    (4)gDir(國外所檔案的實體路徑 ex.\\sin31\FPE_File)
'(3) gbrWebDir(區所檔案的虛擬路徑ex./brp/NPE)    (4)gWebDir(國外所檔案的虛擬路徑 ex./fexp/FPE_File)
IF trim(request("docbranch"))<>empty then
	call getFileServer(trim(request("docbranch")))
	Session("folder_name")="temp/"& trim(request("docbranch")) & "/" & session("folder_name")
Else
	call getFileServer(request("branch"))
End IF	

strtype = Request.QueryString("type")
seq=trim(Request.QueryString("seq"))
draw_file = trim(Request.QueryString("draw_file"))
'file_name = right(draw_file,len(draw_file)-InstrRev(draw_file,"\"))
'Response.Write "filename_flag="& session("filename_flag") & "<BR>"
if session("filename_flag")="source_name" then
	file_name = gbrWebDir &"/"& draw_file
else
	file_name = draw_file
end if
if session("filename_flag")="pic" then
	file_name = session("picturepathN") & "/" & request("folder_name") &"/"& draw_file
else
	btnname = "reg." & trim(request("btnname"))
end if

'Response.Write "filename_flag="& session("filename_flag") & "<BR>"
'Response.Write "file_name="& file_name & "<BR>"
'Response.Write "folder_name="& request("folder_name") & "<BR>"
'Response.Write "seq="& request("seq") & "<BR>"
'Response.Write "seq1="& request("seq1") & "<BR>"
'Response.Write "btnname="& btnname & "<BR>"
'Response.End

Dim strTestFile,objFSO
'strTestFile = Server.MapPath(file_path&"/"&file_name)
'直接使用file_name(傳入完全路徑exp_attach.attach_path)
strTestFile1 = Server.MapPath(file_name)
set objFSO  = CreateObject("Scripting.FileSystemObject")

if chkFileExist_virtual(file_name) = 1 then
	'Response.Write "AAA"
	'將刪除檔案時是將原來word備份起來(日期時間取另一個名字)
	'File_name_new = left(File_name,len(File_name)-4) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) & right (File_name,4)
    File_name_new = left(File_name,len(File_name)-len(right(File_name,len(File_name)-InstrRev(File_name,".")))-1) & "_" & datepart("yyyy",date()) & string(2-len(datepart("m",date())),"0") & datepart("m",date()) & string(2-len(datepart("d",date())),"0") & datepart("d",date()) & string(2-len(hour(time)),"0") & hour(time) & string(2-len(minute(time)),"0") & minute(time) & string(2-len(second(time)),"0") &second(time) &"."& right(File_name,len(File_name)-InstrRev(File_name,"."))
    'Response.Write File_name & "<BR>"
    'response.Write "." & right(File_name,len(File_name)-InstrRev(File_name,".")) & "<BR>"
	'Response.Write "File_name_new="& File_name_new & "<BR>"
	'Response.End
	renameFile1 File_name,File_name_new
	
	'set objFile= objFSO.GetFile(strTestFile1)
	'objFile.Delete
end if
set objFSO=nothing

		
if session("filename_flag")="pic" then
	'改代表圖
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("btbrtdb")
	conn.BeginTrans 
			
	call insert_log_table(conn,"U",prgid,request("tablename"),"seq;seq1",request("seq")&";"&request("seq1"))
	usql = "update "& request("tablename") &" set "
	usql = usql & "pic_file_branch='',pic_file1_path=''"
	usql = usql & " where seq="& request("seq") &" and seq1='"& request("seq1") &"'"
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
	usql = usql & " where seq=" & request("seq") &" and seq1='"& request("seq1") &"' and pic_sqlno='"& request("pic_sqlno") &"'"
	'Response.Write usql &"<Br>"
	conn.Execute usql
				
	usql = "update pic_dmp set "
	usql = usql & "pic_branch='',file1=''"
	usql = usql & " where seq="& request("seq") &" and seq1='"& request("seq1") &"' and pic_sqlno='"& request("pic_sqlno") &"'"
	'Response.Write usql &"<Br>"
	conn.Execute usql
	
	if err.number = 0 and conn.errors.count=0 then
		conn.CommitTrans 
	else
		conn.RollbackTrans 
	end if
end if	
%>

<script language="VBScript">	
	<%if session("filename_flag")="pic" then%>
		window.opener.resetForm
	<%else%>
		if "<%=btnname%>" <> empty then	
			window.opener.<%=btnname%>.disabled = false
		end if
	<%end if%>
	window.close
</script>
