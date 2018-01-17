<%
'新增管制期限檔 ctrl_exp  for 香港標準專利發明案
'H:\_系統檔案\Intranet-brp\系統分析-出專\大陸或英國或歐洲之發明管控香港標準專利案.ppt
function insert_HOctrl_exp(phors_sqlno,phoseq,phoseq1)
	dim i
	for i = 1 to request("HOctrlnum")
		if request("HOctrl_date"&i) <> empty then
			fsql = "insert into ctrl_exp (rs_sqlno,branch,seq,seq1,step_grade,ctrl_type," & _
				"ctrl_remark,ctrl_date,date_ctrl,tran_date,tran_scode) values(" & _
				phors_sqlno & ",'" & session("se_branch") & "'," & phoseq & ",'" & phoseq1 & "',0," & _
				"'" & request("HOctrl_type"&i) & "','" & request("HOctrl_remark"&i) & "'," & _
				chkdatenull(request("HOctrl_date"&i)) & ",'" & request("HOdate_ctrl"&i) & "'," & _
				"getdate(),'" & session("se_scode") & "')"
			'Response.Write "新增管制期限檔-A3-HO法定 table:ctrl_exp <br>" & fsql & "<br>"			
			'Response.End 
			showlog("[<u>Server_savestep.vbs</u>].insert_HOctrl_exp="&fsql)
			conn.execute fsql
		end if		
	next
	'response.end
end function
'新增管制期限檔 ctrl_exp  for 年費法定期限
function insert_a3ctrl_exp(prs_sqlno,pseq,pseq1)
	dim i
	for i = 1 to request("a3ctrlnum")
		if request("a3ctrl_date"&i) <> empty then
			fsql = "insert into ctrl_exp (rs_sqlno,branch,seq,seq1,step_grade,ctrl_type," & _
				"ctrl_remark,ctrl_date,date_ctrl,tran_date,tran_scode) values(" & _
				prs_sqlno & ",'" & session("se_branch") & "'," & pseq & ",'" & pseq1 & "'," & request("nstep_grade") & "," & _
				"'" & request("a3ctrl_type"&i) & "','" & request("a3ctrl_remark"&i) & "'," & _
				chkdatenull(request("a3ctrl_date"&i)) & ",'" & request("a3date_ctrl"&i) & "'," & _
				"getdate(),'" & session("se_scode") & "')"
			'Response.Write "新增管制期限檔-A3年費法定 table:ctrl_exp <br>" & fsql & "<br>"			
			'Response.End 
			showlog("[<u>Server_savestep.vbs</u>].insert_a3ctrl_exp="&fsql)
			conn.execute fsql
		end if		
	next	
end function
'------年費管制處理
function insert_ann_exp(pseq,pseq1,pstep_grade)	
	dim i
	dim actrlnum
	
	set frs = server.CreateObject("ADODB.RECORDSET")	
	
	actrlnum = request("actrlnum")
	yearctrl_flag = request("yearctrl_flag")
	
	'if session("scode")="admin" then
	'	response.write "actrlnum=" & actrlnum &"<BR>"
	'	response.write "yearctrl_flag=" & yearctrl_flag &"<BR>"
	'end if

	if actrlnum <> empty and cstr(actrlnum) <> "0" and yearctrl_flag = "Y" then	
		for i = 1 to actrlnum
			'銷管處理
			if request("harespChk"&i) = "Y" and request("apay_date"&i) <> empty then
				if request("aann_sqlno"&i) <> empty then
					'讀取管制檔資料
					isql = "select * from ann_exp where ann_sqlno='" & request("aann_sqlno"&i) & "'"
					frs.Open isql,conn,1,1
					lmark = frs("mark")
					lann_tran_date = frs("tran_date")
					lann_tran_scode = frs("tran_scode")
					frs.close
					
					if step_date <> empty then
						tresp_date = step_date
					else
						tresp_date = date()
					end if					
										
					'若 ann_sqlno 不為空的時候，表該期限已在管制檔中(ann_imp)，故需刪除
					isql = "insert into ann_resp_exp(ann_sqlno,seq,seq1,pay_times,pay_date,"
					isql = isql & "astep_grade,lstep_grade,mark,add_date,add_step_grade,reset_date,reset_step_grade,"
					isql = isql & "g_date,g_scode,remark,resp_grade,resp_date,resp_type,"
					isql = isql & "ann_tran_date,ann_tran_scode,tran_date,tran_scode)values("
					isql = isql & "'" &request("aann_sqlno"&i)& "'," &pseq& ",'" &pseq1& "',"					
					isql = isql & request("apay_times"&i)& "," &chkdatenull(request("apay_date"&i))& ","
					isql = isql & chknumzero(request("astep_grade"&i)) & "," &chkcharnull(request("lstep_grade"&i))& ","					
					isql = isql & "'" &lmark& "'," 
					if trim(request("aadd_date"&i))<>empty then
						isql = isql & "'" & formatdatetime(request("aadd_date"&i),2) & " " & formatdatetime(request("aadd_date"&i),4) &":"& string(2-len(second(request("aadd_date"&i))),"0") & second(request("aadd_date"&i)) &"',"
					else
						isql = isql & "null,"
					end if
					isql = isql & chknumzero(request("aadd_step_grade"&i))& ","
					if trim(request("areset_date"&i))<>empty then
						isql = isql & "'" & formatdatetime(request("areset_date"&i),2) & " " & formatdatetime(request("areset_date"&i),4) &":"& string(2-len(second(request("areset_date"&i))),"0") & second(request("areset_date"&i)) &"',"
					else
						isql = isql & "null,"
					end if
					isql = isql & chknumzero(request("areset_step_grade"&i))& ","
					if trim(request("ag_date"&i))<>empty then
						isql = isql & "'" & formatdatetime(request("ag_date"&i),2) & " " & formatdatetime(request("ag_date"&i),4) &":"& string(2-len(second(request("ag_date"&i))),"0") & second(request("ag_date"&i)) &"',"
					else
						isql = isql & chkdatenull(request("ag_date"&i)) & ","
					end if
					isql = isql & "'" & request("ag_scode"&i)& "',"
					isql = isql & chkcharnull(request("remark_"&i)) & ","					
					isql = isql & chknumzero(pstep_grade) & "," & chkdatenull(tresp_date) & ","					
					isql = isql & "'" &request("aresp_type"&i)& "','" & formatdatetime(lann_tran_date,2) & " " & formatdatetime(lann_tran_date,4) &":"& string(2-len(second(lann_tran_date)),"0") & second(lann_tran_date) & "',"
					isql = isql & "'" &lann_tran_scode& "',getdate(),'" &session("se_scode")& "')"
					if session("scode")="m983" then
					'Response.Write "寫入年費銷管檔 tabke:ann_resp_exp <br>" & isql & "<br>"
					'response.end
					end if
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.1="&isql)
					conn.execute isql
						
					'刪除年費管制檔 tabke:ann_exp
					isql = "delete from ann_exp where ann_sqlno='" & request("aann_sqlno"&i) & "'"
					'Response.Write "刪除年費管制檔 tabke:ann_imp <br>" & isql & "<br>"
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.2="&isql)
					conn.execute isql
				else
					'若 ann_sqlno 為空的時候，表該期限為一產生使用者就不管制
					'新增年費銷管 Log 檔 table:ann_resp_exp_log
					isql = "insert into ann_resp_exp_log(ud_flag,resp_flag,ud_date,ud_scode,ann_sqlno,seq,seq1,"
					isql = isql & "pay_times,pay_date,astep_grade,lstep_grade,resp_grade,resp_date,resp_type,tran_date,tran_scode,prgid)"
					isql = isql & " values ('A','C',getdate(),'" &session("se_scode")& "',0," & seq & ",'" & seq1 & "',"
					isql = isql & "'" & request("apay_times"&i) & "'," & chkdatenull(request("apay_date"&i)) & ",0,0,'" & pstep_grade & "',"
					isql = isql & chkcharnull(tresp_date) & ",'X',getdate(),'" &session("se_scode")& "','" &prgid& "')"
					
					'Response.Write "新增年費銷管 Log 檔 table:ann_resp_exp_log <br>" & isql & "<br>"
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.3="&isql)
					conn.execute isql
				end if
			end if
		next
		
		'--刪除前,先入檔 ann_exp_log		
		call insert_log_table(conn,"U",prgid,"ann_exp","seq;seq1",pseq&";"&pseq1)		
			
		isql = " delete from ann_exp where seq= '" & pseq & "' and seq1 = '" & pseq1 & "'"
		'Response.Write "刪除期限管制檔 table:ann_exp <br>" & isql & "<br>"			
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.4="&isql)
		conn.execute isql
			
		for i = 1 to actrlnum
			'新增年費管制檔 table:ann_exp
			if request("apay_date"&i) <> empty and request("harespChk"&i) <> "Y" then
				fsql = "insert into ann_exp (seq,seq1,pay_times,pay_date," 
				fsql = fsql & "astep_grade,lstep_grade,add_date,add_step_grade,reset_date,reset_step_grade,"
				fsql = fsql & "remark,g_date,g_scode,tran_date,tran_scode) values("
				fsql = fsql & pseq & ",'" & pseq1 & "',"
				fsql = fsql & "'" & request("apay_times"&i) & "','" & request("apay_date"&i) & "', "
				fsql = fsql & "'" & request("astep_grade"&i) & "','" & request("lstep_grade"&i) & "', "
				if request("submitTask")="U" then
					if trim(request("aadd_date"&i))<>empty then
						fsql = fsql & "'" & formatdatetime(request("aadd_date"&i),2) & " " & formatdatetime(request("aadd_date"&i),4) &":"& string(2-len(second(request("aadd_date"&i))),"0") & second(request("aadd_date"&i)) &"',"
					else
						fsql = fsql & "null,"
					end if
					fsql = fsql & chknumzero(request("aadd_step_grade"&i))& ","
				else
					fsql = fsql & "'" & date() & "'," & pstep_grade & ","
				end if
				if request("reset_flag")="Y" then '控制是否有按過重新產生年費資料
					fsql = fsql & "'" & date() & "'," & pstep_grade & ","
				else
					if trim(request("areset_date"&i))<>empty then
						fsql = fsql & "'" & formatdatetime(request("areset_date"&i),2) & " " & formatdatetime(request("areset_date"&i),4) &":"& string(2-len(second(request("areset_date"&i))),"0") & second(request("areset_date"&i)) &"',"
					else
						fsql = fsql & "null,"
					end if
					fsql = fsql & chknumzero(request("areset_step_grade"&i))& ","
				end if
				fsql = fsql & "'" & request("remark_"&i) & "',"
				if request("reset_flag")="Y" then '控制是否有按過重新產生年費資料
					fsql = fsql & "null,'',"
				else
					if trim(request("ag_date"&i))<>empty then
						fsql = fsql & "'" & formatdatetime(request("ag_date"&i),2) & " " & formatdatetime(request("ag_date"&i),4) &":"& string(2-len(second(request("ag_date"&i))),"0") & second(request("ag_date"&i)) &"',"
					else
						fsql = fsql & "null,"
					end if
					fsql = fsql & "'" & request("ag_scode"&i)& "',"
				end if
				fsql = fsql & "getdate(),'" & session("se_scode") & "')"
				if session("scode")="m983" then
				'Response.Write "新增年費管制期限檔 table:ann_exp <br>" & fsql & "<br>"	
				'response.end
				end if
				showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.5="&fsql)
				conn.execute fsql

				'檢查銷管檔是否有該年度資料，如有要將之紀錄於 Log 檔後刪除，以免日後銷管會 Key 值重複當出來
				isql = "select * from ann_resp_exp where seq = " & pseq & " and seq1 = '" & pseq1 & "' "
				isql = isql & " and pay_times = '" & request("apay_times"&i) & "' "
				frs.Open isql,conn,1,1
				if not frs.EOF then
					'新增年費銷管 Log 檔 table:ann_resp_exp_log
					tkey_field = "seq;seq1;pay_times"
					tkey_value = pseq & ";" & pseq1 & ";" & request("apay_times"&i)
					call insert_log_table(conn,"U",prgid,"ann_resp_exp",tkey_field,tkey_value)
					'刪除年費銷管檔
					isql = "delete from ann_resp_exp where seq = " & pseq & " and seq1 = '" & pseq1 & "' "
					isql = isql & " and pay_times = '" & request("apay_times"&i) & "' "
					'Response.Write "刪除年費銷管檔 table:ann_resp_exp <br>" & fsql & "<br>"	
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.6="&isql)
					conn.execute isql
				end if
				frs.Close
			end if
		next
	end if
	'response.end
	
	if cstr(actrlnum) = "0" and yearctrl_flag = "Y" then
		'--刪除前,先入檔 ann_exp_log		
		call insert_log_table(conn,"D",prgid,"ann_exp","seq;seq1",pseq&";"&pseq1)		
			
		isql = " delete from ann_exp where seq= '" & pseq & "' and seq1 = '" & pseq1 & "'"
		'Response.Write "刪除期限管制檔 table:ann_exp <br>" & isql & "<br>"			
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.7="&isql)
		conn.execute isql
	end if
	
	if trim(request("yearctrl_sqlno")) <> empty and request("yearctrl_sqlno") <> "0" then
		fsql = "select yearctrl_sqlno from exp "
		fsql = fsql & " where seq = " & pseq & " and seq1 = '" & pseq1 & "' "
		frs.open fsql,conn,1,1		
		if (trim(frs("yearctrl_sqlno")) = "0" and request("yearctrl_sqlno") = "") or (trim(frs("yearctrl_sqlno")) = "" and request("yearctrl_sqlno") = "0") then
		else
			insert_exp_rec_log conn,pseq,pseq1,"exp","yearctrl_sqlno",frs("yearctrl_sqlno"),request("yearctrl_sqlno"),prgid,rs_sqlno
			fsql = "update exp set yearctrl_sqlno = " & request("yearctrl_sqlno")
			fsql = fsql & " where seq = " & pseq & " and seq1 = '" & pseq1 & "' "
			'response.write fsql & "<BR>"
			showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.8="&fsql)
			conn.execute fsql			
		end if		
		frs.close		
	end if
	
	set frs = nothing
end function

'銷年費管制入檔 ann_resp_exp
function insert_ann_resp_exp(pseq,pseq1,ppay_times,presp_grade,presp_date,presp_type)
	dim i
	dim fsql
	set frs = server.CreateObject("ADODB.RECORDSET")	
	
	insert_ann_resp_exp = false
	
	fsql = "select * from ann_exp "
	fsql = fsql & " where seq = '" & pseq & "'"
	fsql = fsql & " and seq1 = '" & pseq1 & "'"
	fsql = fsql & " and pay_times = '" & ppay_times &"'"
	'if session("scode") = "admin" and pseq=11408 then
	'response.write fsql & "<br>"
	'response.end	
	'end if
	frs.open fsql,conn,1,1
	if not frs.eof then
		'新增年費銷管檔		
		fsql = "insert into ann_resp_exp(ann_sqlno,seq,seq1,pay_times,pay_date,"
		fsql = fsql & "astep_grade,lstep_grade,mark,add_date,add_step_grade,reset_date,reset_step_grade,"
		fsql = fsql & "g_date,g_scode,remark,resp_grade,resp_date,resp_type,tran_date,tran_scode,"
		fsql = fsql & "ann_tran_date,ann_tran_scode) values("
		fsql = fsql & frs("ann_sqlno") & "," & pseq & ",'" & pseq1 & "',"
		fsql = fsql & ppay_times & ",'" & frs("pay_date") & "',"
		fsql = fsql & "'" & frs("astep_grade") & "','" & frs("lstep_grade") & "',"
		fsql = fsql & "'" & frs("mark") & "'," 
		if trim(frs("add_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("add_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("add_date"),2) & " " & formatdatetime(frs("add_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("add_step_grade") & "'," 
		if trim(frs("reset_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("reset_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("reset_date"),2) & " " & formatdatetime(frs("reset_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("reset_step_grade") & "'," 
		if trim(frs("g_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("g_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("g_date"),2) & " " & formatdatetime(frs("g_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("g_scode") & "','" & frs("remark") & "','" & presp_grade & "',"
		fsql = fsql & chkdatenull(presp_date) & ",'" & presp_type & "',"
		fsql = fsql & "getdate(),'" & session("scode") & "',"
		if trim(frs("tran_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("tran_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("tran_date"),2) & " " & formatdatetime(frs("tran_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("tran_scode") & "')"
		
		'if session("scode") = "admin" then
		'Response.Write "新增年費銷管制檔 table:ann_resp_exp <br>" & fsql & "<br>"
		'response.end	
		'end if
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_resp_exp.1="&fsql)
		conn.execute fsql		
	
		'刪除年費管制檔 tabke:ann_exp
		
		'--刪除前,先入檔 ann_exp_log
		'if session("scode") = "admin" then
		'response.write frs("ann_sqlno") & "<br>"
		'response.end	
		'end if
		call insert_log_table(conn,"D",prgid,"ann_exp","ann_sqlno",frs("ann_sqlno"))
		
		fsql = "delete from ann_exp where ann_sqlno='" & frs("ann_sqlno") & "'"
		'if session("scode") = "admin" then
		'Response.Write "刪除年費管制檔 tabke:ann_exp <br>" & fsql & "<br>"			
		'response.end	
		'end if
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_resp_exp.2="&fsql)
		conn.execute fsql
	end if
	frs.close
		
	if err.number<>0 then
		insert_ann_resp_exp = true
	end if
	set frs = nothing
end function

'寫入年費檔指示進度
function update_ann_step_grade(pseq,pseq1,pspay_times,pepay_times,ptype,pstep_grade)
	dim fsql
	dim i
	
	for i = pspay_times to pepay_times
		call insert_log_table(conn,"U",prgid,"ann_exp","seq;seq1;pay_times",pseq&";"&pseq1&";"&i)
	next
	
	fsql = "update ann_exp "
	fsql = fsql & " set " & lcase(ptype) & "step_grade = " & pstep_grade
	fsql = fsql & " ,tran_date = getdate() "
	fsql = fsql & " ,tran_scode = '" & session("scode") & "' "
	fsql = fsql & " where seq = " & pseq
	fsql = fsql & " and seq1 = '" & pseq1 & "' "
	fsql = fsql & " and pay_times >= '" & pspay_times & "' "
	fsql = fsql & " and pay_times <= '" & pepay_times & "' "

	'Response.Write "修改年費管制檔 table:ann_exp <br>" & fsql & "<br>"
	showlog("[<u>Server_savestep.vbs</u>].update_ann_step_grade="&fsql)
	conn.execute fsql		
	
end function
'新增進度附屬案性檔 step_expd
function insert_step_expd(prs_sqlno) 
	dim i
	for i = 1 to request("codenum")
		if request("rs_class"&i) <> empty and request("rs_code"&i) <> empty and request("act_code"&i) <> empty then
			fsql = "insert into step_expd (rs_sqlno,rs_type,rs_class,rs_code,act_code) values("
			fsql = fsql & prs_sqlno & ",'" & request("rs_type") & "','" & request("rs_class"&i) & "','" & request("rs_code"&i) & "','" & request("act_code"&i) & "')"
			'Response.Write "新增進度附屬案性檔 table:step_expd <br>" & fsql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_step_expd="&fsql)
			conn.execute fsql
		end if
	next
	'response.end
end function
function delete_step_expd(prs_sqlno) 
	fsql = "delete from step_expd where rs_sqlno = '" & prs_sqlno & "'"
	'response.write fsql & "<BR>"
	showlog("[<u>Server_savestep.vbs</u>].delete_step_expd="&fsql)
	conn.execute fsql
end function
'新增管制期限檔 ctrl_exp
function insert_ctrl_exp(prs_sqlno,pseq,pseq1,pstep_grade)	
	dim i
	for i = 1 to request("ctrlnum")
		if request("ctrl_date"&i) <> empty then
			fsql = "insert into ctrl_exp (rs_sqlno,branch,seq,seq1,step_grade,ctrl_type," & _
				"ctrl_remark,ctrl_date,date_ctrl,tran_date,tran_scode) values(" & _
				prs_sqlno & ",'"& session("se_branch") &"'," & pseq & ",'" & pseq1 & "'," & pstep_grade & "," & _
				"'" & request("ctrl_type"&i) & "','" & request("ctrl_remark"&i) & "'," & _
				chkdatenull(request("ctrl_date"&i)) & ",'" & request("date_ctrl"&i) & "'," & _
				"getdate(),'" & session("se_scode") & "')"
				'Response.Write "新增管制期限檔 table:ctrl_exp <br>" & fsql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_ctrl_exp="&fsql)
			conn.execute fsql
		end if		
	next	
	'Response.End 
	
end function
function delete_ctrl_exp(prs_sqlno)
	fsql = "delete from ctrl_exp where rs_sqlno = '" & prs_sqlno & "'"
	if left(prgid,4)="ext2" then
		fsql = fsql & " and substring(ctrl_type,1,1)<>'D' and substring(ctrl_type,1,1)<>'E'"
	else
		fsql = fsql & " and substring(ctrl_type,1,1) in ('D','E')"
	end if
	'response.write fsql&"<br>"
	'response.end
	
	showlog("[<u>Server_savestep.vbs</u>].delete_ctrl_exp="&fsql)
	conn.execute fsql
end function
'銷管制入檔 resp_exp
function insert_resp_exp(prsqlno)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'讀取銷管資料
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'新增管制檔 table:resp_exp
			usql = "insert into resp_exp(ctrl_sqlno,rs_sqlno,branch,seq,seq1,step_grade,"
			usql = usql & "resp_grade,ctrl_type,ctrl_remark,"
			usql = usql & "ctrl_date,date_ctrl,resp_date,resp_type,resp_remark,"
			usql = usql & "Ctrlgs_num,Ctrlgs_sqlno,Back_num,Dctrlgs_num,Dctrlgs_sqlno,Dback_num,"
			usql = usql & "ctrl_tran_date,ctrl_tran_scode,"
			usql = usql & "tran_date,tran_scode)"
			usql = usql & " values('" & rsf("ctrl_sqlno") & "','" & rsf("rs_sqlno") & "','"& session("se_branch") &"',"
			usql = usql & rsf("seq") & ",'" & rsf("seq1") & "','" & rsf("step_grade") & "',"
			usql = usql & "'"& request("nstep_grade") & "',"
			usql = usql & "'" & rsf("ctrl_type") & "','" & rsf("ctrl_remark") & "',"
			usql = usql & "'" & rsf("ctrl_date") & "','" & rsf("date_ctrl") & "','"
			usql = usql & date() & "','G','',"
			usql = usql & chknumzero(rsf("ctrlgs_num")) & "," & chknumzero(rsf("Ctrlgs_sqlno")) & ","
			usql = usql & chknumzero(rsf("Back_num")) & "," & chknumzero(rsf("Dctrlgs_num")) & ","
			usql = usql & chknumzero(rsf("Dctrlgs_sqlno")) & "," & chknumzero(rsf("Dback_num")) & ","
			if rsf("tran_date")<>empty then
				usql = usql & "'" & formatdatetime(rsf("tran_date"),2) & " " & formatdatetime(rsf("tran_date"),4) & "',"
			else
				usql = usql & "null,"
			end if
			usql = usql & "'" & rsf("tran_scode") & "',"
			usql = usql & "getdate(),'" & session("se_scode") & "')"
			
			'Response.Write "新增銷管制檔 table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp.1="&usql)
			conn.execute usql
				
			'刪除期限管制檔 tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "刪除期限管制檔 tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp.2="&usql)
			conn.execute usql
			
			'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
			usql = "update ctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp.3="&usql)
			conn.execute usql
			usql = "update tctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp.4="&usql)
			conn.execute usql
		end if				
		rsf.Close 
	next
	'response.end
end function
function insert_resp_exp2(prsqlno,presp_grade)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'讀取銷管資料
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'新增管制檔 table:resp_exp
			usql = "insert into resp_exp(ctrl_sqlno,rs_sqlno,branch,seq,seq1,step_grade,"
			usql = usql & "resp_grade,ctrl_type,ctrl_remark,"
			usql = usql & "ctrl_date,date_ctrl,resp_date,resp_type,resp_remark,"
			usql = usql & "Ctrlgs_num,Ctrlgs_sqlno,Back_num,Dctrlgs_num,Dctrlgs_sqlno,Dback_num,"
			usql = usql & "ctrl_tran_date,ctrl_tran_scode,"
			usql = usql & "tran_date,tran_scode)"
			usql = usql & " values('" & rsf("ctrl_sqlno") & "','" & rsf("rs_sqlno") & "','"& session("se_branch") &"',"
			usql = usql & rsf("seq") & ",'" & rsf("seq1") & "','" & rsf("step_grade") & "',"
			usql = usql & "'"& presp_grade & "',"
			usql = usql & "'" & rsf("ctrl_type") & "','" & rsf("ctrl_remark") & "',"
			usql = usql & "'" & rsf("ctrl_date") & "','" & rsf("date_ctrl") & "','"
			usql = usql & date() & "','G','',"
			usql = usql & chknumzero(rsf("ctrlgs_num")) & "," & chknumzero(rsf("Ctrlgs_sqlno")) & ","
			usql = usql & chknumzero(rsf("Back_num")) & "," & chknumzero(rsf("Dctrlgs_num")) & ","
			usql = usql & chknumzero(rsf("Dctrlgs_sqlno")) & "," & chknumzero(rsf("Dback_num")) & ","
			usql = usql & "'" & formatdatetime(rsf("tran_date"),2) & " " & formatdatetime(rsf("tran_date"),4) & "',"
			usql = usql & "'" & rsf("tran_scode") & "',"
			usql = usql & "getdate(),'" & session("se_scode") & "')"
			
			'Response.Write "新增銷管制檔 table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp2.1="&usql)
			conn.execute usql
				
			'刪除期限管制檔 tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "刪除期限管制檔 tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp2.2="&usql)
			conn.execute usql
			
			'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
			usql = "update ctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp2.3="&usql)
			conn.execute usql
			usql = "update tctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp2.4="&usql)
			conn.execute usql
		end if				
		rsf.Close 
	next
	'response.end
end function
'銷管制入檔 resp_exp
function insert_resp_exp3(prsqlno,presp_type)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'讀取銷管資料
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'新增管制檔 table:resp_exp
			usql = "insert into resp_exp(ctrl_sqlno,rs_sqlno,branch,seq,seq1,step_grade,"
			usql = usql & "resp_grade,ctrl_type,ctrl_remark,"
			usql = usql & "ctrl_date,date_ctrl,resp_date,resp_type,resp_remark,"
			usql = usql & "Ctrlgs_num,Ctrlgs_sqlno,Back_num,Dctrlgs_num,Dctrlgs_sqlno,Dback_num,"
			usql = usql & "ctrl_tran_date,ctrl_tran_scode,"
			usql = usql & "tran_date,tran_scode)"
			usql = usql & " values('" & rsf("ctrl_sqlno") & "','" & rsf("rs_sqlno") & "','"& session("se_branch") &"',"
			usql = usql & rsf("seq") & ",'" & rsf("seq1") & "','" & rsf("step_grade") & "',"
			usql = usql & "'"& request("nstep_grade") & "',"
			usql = usql & "'" & rsf("ctrl_type") & "','" & rsf("ctrl_remark") & "',"
			usql = usql & "'" & rsf("ctrl_date") & "','" & rsf("date_ctrl") & "','"
			usql = usql & date() & "','"& presp_type &"','',"
			usql = usql & chknumzero(rsf("ctrlgs_num")) & "," & chknumzero(rsf("Ctrlgs_sqlno")) & ","
			usql = usql & chknumzero(rsf("Back_num")) & "," & chknumzero(rsf("Dctrlgs_num")) & ","
			usql = usql & chknumzero(rsf("Dctrlgs_sqlno")) & "," & chknumzero(rsf("Dback_num")) & ","
			if rsf("tran_date")<>empty then
				usql = usql & "'" & formatdatetime(rsf("tran_date"),2) & " " & formatdatetime(rsf("tran_date"),4) & "',"
			else
				usql = usql & "null,"
			end if
			usql = usql & "'" & rsf("tran_scode") & "',"
			usql = usql & "getdate(),'" & session("se_scode") & "')"
			
			'Response.Write "新增銷管制檔 table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp3.1="&usql)
			conn.execute usql
				
			'刪除期限管制檔 tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "刪除期限管制檔 tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp3.2="&usql)
			conn.execute usql
			
			'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
			usql = "update ctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp3.3="&usql)
			conn.execute usql
			usql = "update tctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp3.4="&usql)
			conn.execute usql
		end if				
		rsf.Close 
	next
	'response.end
end function
'銷管制入檔 resp_exp
function insert_resp_exp4(prsqlno,presp_type,presp_remark)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'讀取銷管資料
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'新增管制檔 table:resp_exp
			usql = "insert into resp_exp(ctrl_sqlno,rs_sqlno,branch,seq,seq1,step_grade,"
			usql = usql & "resp_grade,ctrl_type,ctrl_remark,"
			usql = usql & "ctrl_date,date_ctrl,resp_date,resp_type,resp_remark,"
			usql = usql & "Ctrlgs_num,Ctrlgs_sqlno,Back_num,Dctrlgs_num,Dctrlgs_sqlno,Dback_num,"
			usql = usql & "ctrl_tran_date,ctrl_tran_scode,"
			usql = usql & "tran_date,tran_scode)"
			usql = usql & " values('" & rsf("ctrl_sqlno") & "','" & rsf("rs_sqlno") & "','"& session("se_branch") &"',"
			usql = usql & rsf("seq") & ",'" & rsf("seq1") & "','" & rsf("step_grade") & "',"
			usql = usql & "'"& request("nstep_grade") & "',"
			usql = usql & "'" & rsf("ctrl_type") & "','" & rsf("ctrl_remark") & "',"
			usql = usql & "'" & rsf("ctrl_date") & "','" & rsf("date_ctrl") & "','"
			usql = usql & date() & "','"& presp_type &"','" & presp_remark & "',"
			usql = usql & chknumzero(rsf("ctrlgs_num")) & "," & chknumzero(rsf("Ctrlgs_sqlno")) & ","
			usql = usql & chknumzero(rsf("Back_num")) & "," & chknumzero(rsf("Dctrlgs_num")) & ","
			usql = usql & chknumzero(rsf("Dctrlgs_sqlno")) & "," & chknumzero(rsf("Dback_num")) & ","
			if rsf("tran_date")<>empty then
				usql = usql & "'" & formatdatetime(rsf("tran_date"),2) & " " & formatdatetime(rsf("tran_date"),4) & "',"
			else
				usql = usql & "null,"
			end if
			usql = usql & "'" & rsf("tran_scode") & "',"
			usql = usql & "getdate(),'" & session("se_scode") & "')"
			
			'Response.Write "新增銷管制檔 table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp4.1="&usql)
			conn.execute usql
				
			'刪除期限管制檔 tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "刪除期限管制檔 tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp4.2="&usql)
			conn.execute usql
			
			'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
			usql = "update ctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp4.3="&usql)
			conn.execute usql
			usql = "update tctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp4.4="&usql)
			conn.execute usql
		end if				
		rsf.Close 
	next
	'response.end
end function
'程序送件確認時銷管(銷管原因:I:送件確認)
function insert_resp_exp_brconf(prsqlno,presp_grade)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'讀取銷管資料
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'新增管制檔 table:resp_exp
			usql = "insert into resp_exp(ctrl_sqlno,rs_sqlno,branch,seq,seq1,step_grade,"
			usql = usql & "resp_grade,ctrl_type,ctrl_remark,"
			usql = usql & "ctrl_date,date_ctrl,resp_date,resp_type,resp_remark,"
			usql = usql & "Ctrlgs_num,Ctrlgs_sqlno,Back_num,Dctrlgs_num,Dctrlgs_sqlno,Dback_num,"
			usql = usql & "ctrl_tran_date,ctrl_tran_scode,"
			usql = usql & "tran_date,tran_scode)"
			usql = usql & " values('" & rsf("ctrl_sqlno") & "','" & rsf("rs_sqlno") & "','"& session("se_branch") &"',"
			usql = usql & rsf("seq") & ",'" & rsf("seq1") & "','" & rsf("step_grade") & "',"
			usql = usql & "'"& presp_grade & "',"
			usql = usql & "'" & rsf("ctrl_type") & "','" & rsf("ctrl_remark") & "',"
			usql = usql & "'" & rsf("ctrl_date") & "','" & rsf("date_ctrl") & "','"
			usql = usql & date() & "','I','',"
			usql = usql & chknumzero(rsf("ctrlgs_num")) & "," & chknumzero(rsf("Ctrlgs_sqlno")) & ","
			usql = usql & chknumzero(rsf("Back_num")) & "," & chknumzero(rsf("Dctrlgs_num")) & ","
			usql = usql & chknumzero(rsf("Dctrlgs_sqlno")) & "," & chknumzero(rsf("Dback_num")) & ","
			usql = usql & "'" & formatdatetime(rsf("tran_date"),2) & " " & formatdatetime(rsf("tran_date"),4) & "',"
			usql = usql & "'" & rsf("tran_scode") & "',"
			usql = usql & "getdate(),'" & session("se_scode") & "')"
			
			'Response.Write "新增銷管制檔 table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_brconf.1="&usql)
			conn.execute usql
				
			'刪除期限管制檔 tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "刪除期限管制檔 tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_brconf.2="&usql)
			conn.execute usql
			
			'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
			usql = "update ctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_brconf.3="&usql)
			conn.execute usql
			usql = "update tctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_brconf.4="&usql)
			conn.execute usql
		end if				
		rsf.Close 
	next
	'response.end
end function


'銷管制入檔 resp_exp for 結案用
function insert_resp_exp_end(pseq,pseq1)
	set rsf = server.CreateObject("ADODB.RECORDSET")
	fsql = "select * from ctrl_exp where seq="& pseq &" and seq1='"& pseq1 &"'"
	fsql = fsql & " and ctrl_type<>'B8'"
	rsf.open fsql,conn,1,1
	while not rsf.eof
		'新增管制檔 table:resp_exp
		usql = "insert into resp_exp(ctrl_sqlno,rs_sqlno,branch,seq,seq1,step_grade,"
		usql = usql & "resp_grade,ctrl_type,ctrl_remark,"
		usql = usql & "ctrl_date,date_ctrl,resp_date,resp_type,resp_remark,"
		usql = usql & "Ctrlgs_num,Ctrlgs_sqlno,Back_num,Dctrlgs_num,Dctrlgs_sqlno,Dback_num,"
		usql = usql & "ctrl_tran_date,ctrl_tran_scode,"
		usql = usql & "tran_date,tran_scode)"
		usql = usql & " values('" & rsf("ctrl_sqlno") & "','" & rsf("rs_sqlno") & "','"& session("se_branch") &"',"
		usql = usql & rsf("seq") & ",'" & rsf("seq1") & "','" & rsf("step_grade") & "',"
		usql = usql & "'"& request("nstep_grade") & "',"
		usql = usql & "'" & rsf("ctrl_type") & "','" & rsf("ctrl_remark") & "',"
		usql = usql & "'" & rsf("ctrl_date") & "','" & rsf("date_ctrl") & "','"
		usql = usql & date() & "','G','',"
		usql = usql & chknumzero(rsf("ctrlgs_num")) & "," & chknumzero(rsf("Ctrlgs_sqlno")) & ","
		usql = usql & chknumzero(rsf("Back_num")) & "," & chknumzero(rsf("Dctrlgs_num")) & ","
		usql = usql & chknumzero(rsf("Dctrlgs_sqlno")) & "," & chknumzero(rsf("Dback_num")) & ","
		usql = usql & "'" & formatdatetime(rsf("tran_date"),2) & " " & formatdatetime(rsf("tran_date"),4) & "',"
		usql = usql & "'" & rsf("tran_scode") & "',"
		usql = usql & "getdate(),'" & session("se_scode") & "')"
		'Response.Write "新增銷管制檔 table:resp_exp <br>" & usql & "<br>"
		'response.end
		showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_end.1="&usql)
		conn.execute usql
				
		'刪除期限管制檔 tabke:ctrl_exp
		usql = "delete from ctrl_exp where ctrl_sqlno='" & rsf("ctrl_sqlno") & "'"
		'Response.Write "刪除期限管制檔 tabke:ctrl_exp <br>" & usql & "<br>"			
		showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_end.2="&usql)
		conn.execute usql
			
		'銷管制入檔時，同時將稽催回覆檔未回覆的稽催給X不需處理
		usql = "update ctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
		showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_end.3="&usql)
		conn.execute usql
		usql = "update tctrlgs_exp set back_flag='X' where ctrl_sqlno='"& rsf("ctrl_sqlno") &"'"
		showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_end.4="&usql)
		conn.execute usql
		rsf.movenext	
	wend
	rsf.Close 
end function

'銷年費管制入檔 ann_resp_exp
function insert_ann_resp_exp_end(pseq,pseq1,presp_grade,presp_date,presp_type)
	dim i
	dim fsql
	set frs = server.CreateObject("ADODB.RECORDSET")	
	
	insert_ann_resp_exp_end = false
	
	fsql = "select * from ann_exp "
	fsql = fsql & " where seq = " & pseq
	fsql = fsql & " and seq1 = '" & pseq1 & "' "
	frs.open fsql,conn,1,1
	while not frs.eof
		'新增年費銷管檔		
		fsql = "insert into ann_resp_exp(ann_sqlno,seq,seq1,pay_times,pay_date,"
		fsql = fsql & "astep_grade,lstep_grade,mark,add_date,add_step_grade,reset_date,reset_step_grade,"
		fsql = fsql & "g_date,g_scode,remark,resp_grade,resp_date,resp_type,tran_date,tran_scode,"
		fsql = fsql & "ann_tran_date,ann_tran_scode) values("
		fsql = fsql & frs("ann_sqlno") & "," & pseq & ",'" & pseq1 & "',"
		fsql = fsql & frs("pay_times") & ",'" & frs("pay_date") & "',"
		fsql = fsql & "'" & frs("astep_grade") & "','" & frs("lstep_grade") & "',"
		fsql = fsql & "'" & frs("mark") & "'," 
		if trim(frs("add_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("add_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("add_date"),2) & " " & formatdatetime(frs("add_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("add_step_grade") & "'," 
		if trim(frs("reset_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("reset_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("reset_date"),2) & " " & formatdatetime(frs("reset_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("reset_step_grade") & "'," 
		if trim(frs("g_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("g_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("g_date"),2) & " " & formatdatetime(frs("g_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("g_scode") & "','" & frs("remark") & "','" & presp_grade & "',"
		fsql = fsql & chkdatenull(presp_date) & ",'" & presp_type & "',"
		fsql = fsql & "getdate(),'" & session("scode") & "',"
		if trim(frs("tran_date"))<>empty then 
			fsql = fsql & chkdatenull2(frs("tran_date")) &","
			'fsql = fsql & "'" & formatdatetime(frs("tran_date"),2) & " " & formatdatetime(frs("tran_date"),4) & "'," 
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & "'" & frs("tran_scode") & "')"
		
		'Response.Write "新增年費銷管制檔 table:ann_resp_exp <br>" & fsql & "<br>"
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_resp_exp_end.1="&fsql)
		conn.execute fsql
	
		'刪除年費管制檔 tabke:ann_exp
		
		'--刪除前,先入檔 ann_exp_log
		call insert_log_table(conn,"D",prgid,"ann_exp","ann_sqlno",frs("ann_sqlno"))
		frs.movenext
	wend
	frs.close
	fsql = "delete from ann_exp"
	fsql = fsql & " where seq = " & pseq
	fsql = fsql & " and seq1 = '" & pseq1 & "' "
	'Response.Write "刪除年費管制檔 tabke:ann_exp <br>" & fsql & "<br>"			
	showlog("[<u>Server_savestep.vbs</u>].insert_ann_resp_exp_end.2="&fsql)
	conn.execute fsql
		
	if err.number<>0 then
		insert_ann_resp_exp_end = true
	end if
	set frs = nothing
end function

function delete_resp_exp(prs_sqlno)
	dim fsql
	fsql = "delete from resp_exp where rs_sqlno = '" & prs_sqlno & "'"
	showlog("[<u>Server_savestep.vbs</u>].delete_resp_exp="&fsql)
	conn.execute fsql
end function
'入結案異動檔
function insert_exp_end(pend_flag,pback_flag)
	dim fsql
	'dim isql 
	'set rsf = server.CreateObject("ADODB.RECORDSET")
	'isql = "select * from case_exp where case_no='"& case_no &"'"
	if pend_flag = "Y" then
		end_flag = "end"
	elseif pback_flag = "Y" then
		end_flag = "back"
	end if
	fsql = "insert into exp_end (case_no,in_no,rs_sqlno,seq,seq1,step_grade,end_flag,"
	fsql = fsql & "end_code,endremark,br_reason,in_scode,in_date,tran_date,tran_scode"
	fsql = fsql & ") values(" 
	fsql = fsql & "'" & case_no & "','"& in_no &"',"& rs_sqlno &"," & seq & ",'" & seq1 & "',"
	fsql = fsql & request("nstep_grade") & "," & chkdatenull(end_flag) & ","
	fsql = fsql & chkdatenull(request("step_end_code")) & "," & chkdatenull(request("step_endremark")) & ","
	fsql = fsql & chkdatenull(request("flag_remark")) & ","
	fsql = fsql & "'" & session("scode") & "',getdate(),"
	fsql = fsql & "getdate(),'" & session("scode") & "'"
	fsql = fsql & ")"
	'Response.Write "新增結案異動檔 table:exp_end <br>" & fsql & "<br>"			
	'Response.End 
	showlog("[<u>Server_savestep.vbs</u>].insert_exp_end="&fsql)
	conn.execute fsql

end function
'掃描文件新增至 exp_attach
function insert_exp_scan(pexp_sqlno,pseq,pseq1,pstep_grade,pscan_path,pattach_no,pchk_status)
	dim fsql
	dim tscan_name
	
	ar_scan = split(pscan_path,"/")
	tscan_name = ar_scan(ubound(ar_scan))
	
	fsql = "insert into exp_attach (seq,seq1,step_grade,exp_sqlno,source,in_date,in_scode"
	fsql = fsql & ",attach_no,attach_path,attach_desc,attach_name,source_name"
	fsql = fsql & ",attach_flag,chk_status,mark,open_flag,tran_date,tran_scode)values("
	fsql = fsql & pseq & ",'" & pseq1 & "','" & pstep_grade &"','"& pexp_sqlno & "','SCAN','" & date() & "' "
	fsql = fsql & ",'" & session("scode") & "','" & pattach_no & "','" & pscan_path & "' "
	fsql = fsql & ",'掃描文件','" & tscan_name & "','" & tscan_name & "' "
	fsql = fsql & ",'A',"
	if pchk_status = "Y1" then
		fsql = fsql & "'Y',"
	else
		fsql = fsql & "'N',"
	end if
	fsql = fsql & "'N','N',getdate(),'" & session("scode") & "') "

	'Response.Write "新增掃描文件檔 table:exp_attach <br>" & fsql & "<br>"
	showlog("[<u>Server_savestep.vbs</u>].insert_exp_scan="&fsql)
	conn.execute fsql		
end function
%>
