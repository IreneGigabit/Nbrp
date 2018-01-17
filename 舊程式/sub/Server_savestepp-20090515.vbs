<%
'=================  國內案  ================'
'ctrl_dmp , resp_dmp , dmp_attach
'處理上傳圖檔的部份
Function upin_dmp_attach_for_job(pseq,pseq1,pstep_grade,pjob_branch,pjob_sqlno)
	dim i
	'目前資料庫中有的最大值
	maxAttach_no = request("uploadfield")&"_maxAttach_no"
	'response.write  "maxAttach_no=" & maxAttach_no &"<br>"
	'目前畫面上的最大值
	filenum=request("uploadfield")&"filenum"
	'response.write  "filenum=" & filenum &"<br>"
	'response.write  "filenum1=" & request(filenum) &"<br>"
	sqlnum=request("uploadfield")&"sqlnum"
	'response.write  "sqlnum=" & sqlnum &"<br>"
	'目前table的筆數
	attach_cnt=request("uploadfield")&"_attach_cnt"
	'response.write "attach_cnt="& attach_cnt &"<BR>"
	'欄位名稱
	uploadfield = trim(request("uploadfield"))
	'response.write "uploadfield="& uploadfield & "<BR>"
	'response.write "maxattach_no="& request(maxattach_no) &"<BR>"
	
	uploadsource = trim(request("uploadsource"))
	
	for i=1 to cint(request(sqlnum))
		'response.write "upload-dbflag:"& trim(request(uploadfield & "_dbflag" & i)) &"<BR>"
		'response.write trim(request(uploadfield & "_exp_sqlno" & i)) &"<BR>"
		select case trim(request(uploadfield & "_dbflag" & i))
			case "A"
				'當上傳路徑不為空的 and attach_sqlno為空的,才需要新增
				IF trim(request(uploadfield&i))<>empty and trim(request(uploadfield & "_attach_sqlno" & i)) ="" then
						IF trim(request(uploadfield & "_open_flag" & i))=empty then
							topen_flag="N"
						Else	
							topen_flag="Y"
						End IF
						fsql = "insert into dmp_attach (dmp_sqlno,Seq,seq1,step_grade,case_no,job_branch,job_sqlno,Source"
						fsql = fsql & ",in_date,in_scode,Attach_no,attach_path,doc_type,attach_desc,Attach_name"
						fsql = fsql & ",source_name,Attach_size,attach_page,attach_flag,attach_flagbr,Mark,open_flag"
						fsql = fsql & ",tran_date,tran_scode,in_no"
						If lcase(prgid)="dmp3a2" then
						 	fsql = fsql & ",tran_datef,tran_scodef"
						End IF
						fSQL = fsql & ") values ("
						fsql = fsql & "'"& trim(request(uploadfield & "_dmp_sqlno" & i)) &"','"& trim(pseq) &"','"& trim(pseq1) &"',"
						fsql = fsql & "'"& trim(pstep_grade) &"','"& trim(request(uploadfield & "_case_no" & i)) &"','"& trim(session("se_branch")) &"','"& trim(pjob_sqlno) &"',"
						fsql = fsql & "'"& uploadsource &"',getdate(),'"& session("scode")&"','"& trim(request(uploadfield & "_attach_no" & i)) &"','"& request(uploadfield&i) &"',"
						fsql = fsql & "'" & trim(request(uploadfield & "_temp_doc" & i)) & "',"
						fsql = fsql & "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"',"
						fsql = fsql & "'" & trim(request(uploadfield & "_source_name" & i)) & "','"& trim(request(uploadfield & "_size" & i)) &"',"
						fsql = fsql & "'" & trim(request(uploadfield & "_page" & i)) & "',"
						fsql = fsql & "'A','','','"& trim(topen_flag) &"',getdate(),'"& session("scode") & "'"
						fsql = fsql & ",'"& request("in_no") &"'"
						If lcase(prgid)="dmp3a2" then
						 	fsql = fsql & ",getdate(),'"& session("scode") &"'"
						End IF
						fsql = fsql & ")"
					'	if session("scode")="admin" then
					'		response.write "資料庫無資料新增dmp_attach"& i &"<br>" &fsql&"<br><br>"
					'		Response.End 
					'	end if
						attachno = attachno + 1
						Conn.execute fsql
				end if	
			case "U"
					'當attach_sqlno <> empty時 , 而且上傳的路徑又是空的時候,表示要刪除該筆資料,而非修改
					if request(uploadfield&"_attach_sqlno"&i) <> empty and trim(request(uploadfield&i)) = empty then
						call insert_log_table(conn,"D",prgid,"dmp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
				
						'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
						if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
							dsql = "update dmp_attach set attach_flag='D'"
							If lcase(prgid)="dmp3a2" then
							 	dsql = dsql & ",tran_datef=getdate(),tran_scodef='"& session("scode") &"'"
							End IF
							dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
							'response.write "delete dmp_attach "& i &"="&dsql&"<br><br>"
							Conn.Execute dsql
						else
							'不需要處理,表示原本db就沒有值
						end if
					else
						call insert_log_table(conn,"U",prgid,"dmp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
						IF trim(request(uploadfield & "_open_flag" & i))=empty then
							topen_flag="N"
						Else	
							topen_flag="Y"
						End IF
						uSQL = "Update dmp_attach set Source='"& uploadsource &"'"
						uSQL = uSQL & ",attach_path='"& request(uploadfield&i) &"'"
						uSQL = uSQL & ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" 
						uSQL = uSQL & ",attach_name='"& request(uploadfield & "_name" & i) &"'"
						uSQL = uSQL & ",attach_size='"& request(uploadfield & "_size" & i) &"'"
						uSQL = uSQL & ",attach_page='"& request(uploadfield & "_page" & i) &"'"
						uSQL = uSQL & ",source_name='"& request(uploadfield & "_source_name" & i) &"'"
						uSQL = uSQL & ",doc_type='"& request(uploadfield & "_temp_doc" & i) &"'"
						uSQL = uSQL & ",attach_flag='U'"
						uSQL = uSQL & ",open_flag='"& topen_flag &"'"
						uSQL = uSQL & ",tran_date=getdate()"
						uSQL = uSQL & ",tran_scode='"&  session("scode") &"'"
						If lcase(prgid)="dmp3a2" then
							uSQL = uSQL & ",tran_datef=getdate()"
							uSQL = uSQL & ",tran_scodef='"&  session("scode") &"'"
						End IF
						uSQL = uSQL & " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
						'response.write "更新資料 < Update dmp_attach"& i &"=" & uSQL & "<br><br>"
						'response.end
						Conn.execute uSQL
					end if
			
			case "D"
				call insert_log_table(conn,"D",prgid,"dmp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))

				'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
				if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
					dsql = "update dmp_attach set attach_flag='D'"
					If lcase(prgid)="dmp3a2" then
					 	fsql = fsql & ",tran_datef=getdate(),tran_scodef='"& session("scode") &"'"
					End IF
					dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "delete dmp_attach "& i &"="&dsql&"<br><br>"
					Conn.Execute dsql
				else
					'不需要處理,表示原本db就沒有值
				end if
		end select	
	next
	'response.end
End Function
'處理上傳圖檔的部份_將文入附件入到各區所
Function upin_dmp_attach_for_branch(pconn,pbranch,pseq,pseq1,pstep_grade,prs_sqlno,pjob_sqlno,pbr_sqlno)
	Call getFileServer(pbranch)
	dim i
	'目前資料庫中有的最大值
	maxAttach_no = request("uploadfield")&"_maxAttach_no"
	'response.write  "maxAttach_no=" & maxAttach_no &"<br>"
	'目前畫面上的最大值
	filenum=request("uploadfield")&"filenum"
	'response.write  "filenum=" & filenum &"<br>"
	'response.write  "filenum1=" & request(filenum) &"<br>"
	sqlnum=request("uploadfield")&"sqlnum"
	'response.write  "sqlnum=" & sqlnum &"<br>"
	'目前table的筆數
	attach_cnt=request("uploadfield")&"_attach_cnt"
	'response.write "attach_cnt="& attach_cnt &"<BR>"
	'欄位名稱
	uploadfield = trim(request("uploadfield"))
	'response.write "uploadfield="& uploadfield & "<BR>"
	'response.write "maxattach_no="& request(maxattach_no) &"<BR>"
	
	uploadsource = trim(request("uploadsource"))
	
	for i=1 to cint(request(sqlnum))
		select case trim(request(uploadfield & "_dbflag" & i))
			case "A"
				'當上傳路徑不為空的 and attach_sqlno為空的,才需要新增
				IF trim(request(uploadfield&i))<>empty and trim(request(uploadfield & "_attach_sqlno" & i)) ="" then
					IF trim(request(uploadfield & "_open_flag" & i))=empty then
						topen_flag="N"
					Else	
						topen_flag="Y"
					End IF
					attach_path=replace(trim(request(uploadfield&i)),"/temp/"&pbranch,"")
					fsql = "insert into dmp_attach (dmp_sqlno,Seq,seq1,step_grade,rs_sqlno,case_no,job_branch,br_sqlno,Source"
					fsql = fsql & ",in_date,in_scode,Attach_no,attach_path,doc_type,attach_desc,Attach_name"
					fsql = fsql & ",source_name,Attach_size,attach_page,attach_flag,attach_flagbr,Mark,open_flag"
					fsql = fsql & ",tran_date,tran_scode,in_no"
					If lcase(prgid)="dmp3a2" then
					 	fsql = fsql & ",tran_datef,tran_scodef"
					End IF
					fSQL = fsql & ") values ("
					fsql = fsql & "'"& trim(request(uploadfield & "_dmp_sqlno" & i)) &"','"& trim(pseq) &"','"& trim(pseq1) &"',"
					fsql = fsql & "'"& trim(pstep_grade) &"','"&trim(prs_sqlno)&"','"& trim(request(uploadfield & "_case_no" & i)) &"','"& trim(session("se_branch")) &"','"& trim(pbr_sqlno) &"',"
					fsql = fsql & "'"& uploadsource &"',getdate(),'"& session("scode")&"','"& trim(request(uploadfield & "_attach_no" & i)) &"','"& attach_path &"',"
					fsql = fsql & "'" & trim(request(uploadfield & "_temp_doc" & i)) & "',"
					fsql = fsql & "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"',"
					fsql = fsql & "'" & trim(request(uploadfield & "_source_name" & i)) & "','"& trim(request(uploadfield & "_size" & i)) &"',"
					fsql = fsql & "'" & trim(request(uploadfield & "_page" & i)) & "',"
					fsql = fsql & "'A','N','','"& trim(topen_flag) &"',getdate(),'"& session("scode") & "'"
					fsql = fsql & ",'"& request("in_no") &"'"
					If lcase(prgid)="dmp3a2" then
					 	fsql = fsql & ",getdate(),'"& session("scode") &"'"
					End IF
					fsql = fsql & ")"
					'response.write "資料庫無資料新增dmp_attach"& i &"<br>" &fsql&"<br><br>"
					attachno = attachno + 1
					pConn.execute fsql
					if err.number=0  then
						Call copyfile_tobranch(trim(request("branch")),trim(request(uploadfield&i))) 
					End IF	
				end if	
			case "U"
				'當attach_sqlno <> empty時 , 而且上傳的路徑又是空的時候,表示要刪除該筆資料,而非修改
				if request(uploadfield&"_attach_sqlno"&i) <> empty and trim(request(uploadfield&i)) = empty then
					call insert_log_table(conn,"D",prgid,"dmp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
					
					'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
					if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
						dsql = "update dmp_attach set attach_flag='D'"
						dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
						'response.write "delete dmp_attach "& i &"="&dsql&"<br><br>"
						pConn.Execute dsql
					else
						'不需要處理,表示原本db就沒有值
					end if
				else
					call insert_log_table(pconn,"U",prgid,"dmp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
					IF trim(request(uploadfield & "_open_flag" & i))=empty then
						topen_flag="N"
					Else	
						topen_flag="Y"
					End IF
					attach_path=replace(trim(request(uploadfield&i)),"/temp/"&pbranch,"")
							
					uSQL = "Update dmp_attach set Source='"& uploadsource &"'"
					uSQL = uSQL & ",attach_path='"& attach_path &"'"
					uSQL = uSQL & ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" 
					uSQL = uSQL & ",attach_name='"& request(uploadfield & "_name" & i) &"'"
					uSQL = uSQL & ",attach_size='"& request(uploadfield & "_size" & i) &"'"
					uSQL = uSQL & ",attach_page='"& request(uploadfield & "_page" & i) &"'"
					uSQL = uSQL & ",source_name='"& request(uploadfield & "_source_name" & i) &"'"
					uSQL = uSQL & ",doc_type='"& request(uploadfield & "_temp_doc" & i) &"'"
					uSQL = uSQL & ",attach_flag='U',job_branch='"& pbranch &"',br_sqlno="& pbr_sqlno
					uSQL = uSQL & ",attach_flagbr='N',open_flag='"& topen_flag &"'"
					uSQL = uSQL & ",tran_date=getdate()"
					uSQL = uSQL & ",tran_scode='"&  session("scode") &"'"
					If lcase(prgid)="dmp3a2" then
						uSQL = uSQL & ",tran_datef=getdate()"
						uSQL = uSQL & ",tran_scodef='"&  session("scode") &"'"
					End IF
					uSQL = uSQL & " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "更新資料 < Update dmp_attach"& i &"=" & uSQL & "<br><br>"
					'response.end
					pConn.execute uSQL
					if err.number=0  then
						'response.write trim(request(uploadfield&i)) &"<BR>"
						'response.end
						Call copyfile_tobranch(trim(request("branch")),trim(request(uploadfield&i))) 
					End IF	
				end if
			case "D"
				call insert_log_table(pconn,"D",prgid,"dmp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
				'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
				if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
					dsql = "update dmp_attach set attach_flag='D'"
					dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "delete dmp_attach "& i &"="&dsql&"<br><br>"
					pConn.Execute dsql
				else
					'不需要處理,表示原本db就沒有值
				end if
		end select	
	next
End Function
'新增管制期限檔 ctrl_dmp
function insert_ctrl_dmp(prs_sqlno,pseq,pseq1,pstep_grade)	
	dim i
	dim fsql
	for i=1 to request("ctrlnum")
		'Response.Write "管制:" & i & "--" & request("delchk"&i) & "<BR>"
		if request("ctrl_type"&i)<>empty and trim(request("ctrl_date"&i))<>empty then
			fsql = "insert into ctrl_dmp(step_sqlno,branch,seq,seq1,step_grade,ctrl_type,ctrl_remark,ctrl_date,tran_date,tran_scode)"
			fsql = fsql & " values("& prs_sqlno &",'" & session("se_branch") & "'," & pseq & ","
			fsql = fsql & "'" & trim(pseq1) & "',"
			fsql = fsql & pstep_grade & ",'" & request("ctrl_type"&i) & "'," & chkcharnull2(request("ctrl_remark"&i)) & ","
			fsql = fsql & chkdatenull(request("ctrl_date"&i)) & ",getdate(),'" & session("se_scode") & "')"
			'response.write fsql & "<BR>"
			conn.execute(fsql)
			if err.number<>0 then geterrmsg
		end if
	next
end function
'修改管制期限檔
function update_ctrl_dmp(prs_sqlno,pseq,pseq1,pstep_grade)	
	i = 1
	for i=1 to request("ctrlnum")
		if request("delchk"&i)=false and request("io_flg"&i)="Y" then
			if request("ctrl_type"&i)<>empty and trim(request("ctrl_date"&i))<>empty then
				sql = "insert into ctrl_dmp(step_sqlno,branch,seq,seq1,step_grade,ctrl_type,ctrl_remark,ctrl_date,tran_date,tran_scode)"
				sql = sql & " values(" & prs_sqlno & ",'" & session("se_branch") & "'," & pseq & ",'" & trim(pseq1) & "',"
				sql = sql & pstep_grade & ",'" & request("ctrl_type"&i) & "'," & chkcharnull2(request("ctrl_remark"&i)) & ","
				sql = sql & chkdatenull(formatdatetime(request("ctrl_date"&i),2)) & ",getdate(),'" & session("se_scode") & "')"
				conn.execute(sql)
				if err.number<>0 then geterrmsg
				'Response.Write "更新ctrl_dmp" & i & "--" & "<br>" & sql & "<BR>"
			end if
		end if
	next
end function
'刪除管制期限檔
function delete_ctrl_dmp(prs_sqlno)
	fsql = "delete from ctrl_dmp where step_sqlno = '" & prs_sqlno & "'"
	'response.write fsql&"<br>"
	'response.end
	conn.execute fsql
	if err.number<>0 then geterrmsg
end function
'銷管制入檔 resp_dmp
function insert_resp_dmp(prsqlno)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")

	'銷管制入檔
	if request("rsqlno") <> empty then
		ar = split(request("rsqlno"),";")
		for i = 0 to ubound(ar) -1
			'讀取銷管資料
			isql = "select * from ctrl_dmp where sqlno='" & ar(i) & "'"
			rsf.Open isql,conn,1,1
			if not rsf.EOF then
				'新增至 resp_dmp
				sql = "insert into resp_dmp(sqlno,step_sqlno,branch,seq,seq1,step_grade,resp_grade,"
				sql = sql & "ctrl_type,ctrl_remark,ctrl_date,resp_date,tran_date,tran_scode)"
				sql = sql & " values('" & rsf("sqlno") & "'," & rsf("step_sqlno") & ","
				sql = sql & "'" & rsf("branch") & "'," & rsf("seq") & ","
				if prgid="brpa24" then
					sql = sql & "'" & trim(request("grnseq1")) & "',"
				else
					sql = sql & "'" & trim(request("seq1")) & "',"
				end if
				sql = sql & "'" & rsf("step_grade") & "','" & request("nstep_grade") & "','" & rsf("ctrl_type") & "',"
				sql = sql & "'" & rsf("ctrl_remark") & "','" & rsf("ctrl_date") & "','" & request("step_date") & "',getdate(),'" & session("se_scode") & "')"
				conn.execute(sql)
				if err.number<>0 then geterrmsg
				'由 ctrl_dmp 中刪除
				sql = "delete from ctrl_dmp where sqlno='" & ar(i) & "'"
				conn.execute(sql)
				if err.number<>0 then geterrmsg
				'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
				usql = "update ctrlgs_mgp set back_flag='X' where ctrl_sqlno='"& ar(i) &"'"
			'	conn.execute usql
			end if
			rsf.Close 
		next
	end if	
end function
'刪除銷管制
function delete_resp_dmp(prs_sqlno)
	dim fsql
	fsql = "delete from resp_dmp where step_sqlno = '" & prs_sqlno & "'"
	conn.execute fsql
end function
'掃描文件新增至 dmp_attach
function insert_dmp_scan(pdmp_sqlno,pseq,pseq1,pstep_grade,pscan_path,pattach_no,pchk_status)
	dim fsql
	dim tscan_name
	
	ar_scan = split(pscan_path,"/")
	tscan_name = ar_scan(ubound(ar_scan))
	
	fsql = "insert into dmp_attach (seq,seq1,step_grade,dmp_sqlno,source,in_date,in_scode"
	fsql = fsql & ",attach_no,attach_path,attach_desc,attach_name,source_name"
	fsql = fsql & ",attach_flag,chk_status,mark,tran_date,tran_scode)values("
	fsql = fsql & pseq & ",'" & pseq1 & "','" & pstep_grade &"','"& pdmp_sqlno & "','SCAN','" & date() & "' "
	fsql = fsql & ",'" & session("scode") & "','" & pattach_no & "','" & pscan_path & "' "
	fsql = fsql & ",'掃描文件','" & tscan_name & "','" & tscan_name & "' "
	fsql = fsql & ",'A','" & pchk_status & "',"
	if pchk_status = "Y1" then
		fsql = fsql & "'Y',"
	else
		fsql = fsql & "'N',"
	end if
	fsql = fsql & "getdate(),'" & session("scode") & "') "

	'Response.Write "新增掃描文件檔 table:dmp_attach <br>" & fsql & "<br>"
	conn.execute fsql		
end function

'修改掃描文件資料
function update_dmp_scan(pseq,pseq1,pstep_grade,pscan_path,pattach_no)
	dim fsql
	dim tscan_name
	
	call insert_log_table(conn,"U",prgid,"dmp_attach","seq;seq1;step_grade;source;attach_no",pseq&";"&pseq1&";"&pstep_grade&";scan;1")
	
	ar_scan = split(pscan_path,"/")
	tscan_name = ar_scan(ubound(ar_scan))
	
	fsql = "update dmp_attach "
	fsql = fsql & " set attach_path = '" & pscan_path & "' "
	fsql = fsql & " ,attach_name = '" & tscan_name & "' "
	fsql = fsql & " ,source_name = '" & tscan_name & "' "
	fsql = fsql & " ,attach_flag = 'U' "
	fsql = fsql & " ,tran_date = getdate() "
	fsql = fsql & " ,tran_scode = '" & session("scode") & "' "
	fsql = fsql & " where seq = " & pseq
	fsql = fsql & "   and seq1 = '" & pseq1 & "' "
	fsql = fsql & "   and step_grade = '" & pstep_grade & "' "
	fsql = fsql & "   and source = 'SCAN' "
	fsql = fsql & "   and attach_no = '" & pattach_no & "' "
	
	Response.Write "修改掃描文件檔 table:dmp_attach <br>" & fsql & "<br>"
	conn.execute fsql		
end function

'刪除掃描文件資料
function delete_dmp_scan(pseq,pseq1,pstep_grade,pscan_path,pattach_no)
	dim fsql
	
	call insert_log_table(conn,"D",prgid,"dmp_attach","seq;seq1;step_grade;source;attach_no",pseq&";"&pseq1&";"&pstep_grade&";scan;1")
	
	fsql = "update dmp_attach "
	fsql = fsql & " set attach_flag = 'D' "
	fsql = fsql & " ,tran_date = getdate() "
	fsql = fsql & " ,tran_scode = '" & session("scode") & "' "
	fsql = fsql & " where seq = " & pseq
	fsql = fsql & "   and seq1 = '" & pseq1 & "' "
	fsql = fsql & "   and step_grade = '" & pstep_grade & "' "
	
	Response.Write "刪除進度文件檔 table:dmp_attach <br>" & fsql & "<br>"
	conn.execute fsql		
end function
%>
