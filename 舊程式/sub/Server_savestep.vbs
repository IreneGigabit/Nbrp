<%
'�s�W�ި������ ctrl_exp  for ����зǱM�Q�o����
'H:\_�t���ɮ�\Intranet-brp\�t�Τ��R-�X�M\�j���έ^��μڬw���o���ޱ�����зǱM�Q��.ppt
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
			'Response.Write "�s�W�ި������-A3-HO�k�w table:ctrl_exp <br>" & fsql & "<br>"			
			'Response.End 
			showlog("[<u>Server_savestep.vbs</u>].insert_HOctrl_exp="&fsql)
			conn.execute fsql
		end if		
	next
	'response.end
end function
'�s�W�ި������ ctrl_exp  for �~�O�k�w����
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
			'Response.Write "�s�W�ި������-A3�~�O�k�w table:ctrl_exp <br>" & fsql & "<br>"			
			'Response.End 
			showlog("[<u>Server_savestep.vbs</u>].insert_a3ctrl_exp="&fsql)
			conn.execute fsql
		end if		
	next	
end function
'------�~�O�ި�B�z
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
			'�P�޳B�z
			if request("harespChk"&i) = "Y" and request("apay_date"&i) <> empty then
				if request("aann_sqlno"&i) <> empty then
					'Ū���ި��ɸ��
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
										
					'�Y ann_sqlno �����Ū��ɭԡA��Ӵ����w�b�ި��ɤ�(ann_imp)�A�G�ݧR��
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
					'Response.Write "�g�J�~�O�P���� tabke:ann_resp_exp <br>" & isql & "<br>"
					'response.end
					end if
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.1="&isql)
					conn.execute isql
						
					'�R���~�O�ި��� tabke:ann_exp
					isql = "delete from ann_exp where ann_sqlno='" & request("aann_sqlno"&i) & "'"
					'Response.Write "�R���~�O�ި��� tabke:ann_imp <br>" & isql & "<br>"
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.2="&isql)
					conn.execute isql
				else
					'�Y ann_sqlno ���Ū��ɭԡA��Ӵ������@���ͨϥΪ̴N���ި�
					'�s�W�~�O�P�� Log �� table:ann_resp_exp_log
					isql = "insert into ann_resp_exp_log(ud_flag,resp_flag,ud_date,ud_scode,ann_sqlno,seq,seq1,"
					isql = isql & "pay_times,pay_date,astep_grade,lstep_grade,resp_grade,resp_date,resp_type,tran_date,tran_scode,prgid)"
					isql = isql & " values ('A','C',getdate(),'" &session("se_scode")& "',0," & seq & ",'" & seq1 & "',"
					isql = isql & "'" & request("apay_times"&i) & "'," & chkdatenull(request("apay_date"&i)) & ",0,0,'" & pstep_grade & "',"
					isql = isql & chkcharnull(tresp_date) & ",'X',getdate(),'" &session("se_scode")& "','" &prgid& "')"
					
					'Response.Write "�s�W�~�O�P�� Log �� table:ann_resp_exp_log <br>" & isql & "<br>"
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.3="&isql)
					conn.execute isql
				end if
			end if
		next
		
		'--�R���e,���J�� ann_exp_log		
		call insert_log_table(conn,"U",prgid,"ann_exp","seq;seq1",pseq&";"&pseq1)		
			
		isql = " delete from ann_exp where seq= '" & pseq & "' and seq1 = '" & pseq1 & "'"
		'Response.Write "�R�������ި��� table:ann_exp <br>" & isql & "<br>"			
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.4="&isql)
		conn.execute isql
			
		for i = 1 to actrlnum
			'�s�W�~�O�ި��� table:ann_exp
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
				if request("reset_flag")="Y" then '����O�_�����L���s���ͦ~�O���
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
				if request("reset_flag")="Y" then '����O�_�����L���s���ͦ~�O���
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
				'Response.Write "�s�W�~�O�ި������ table:ann_exp <br>" & fsql & "<br>"	
				'response.end
				end if
				showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.5="&fsql)
				conn.execute fsql

				'�ˬd�P���ɬO�_���Ӧ~�׸�ơA�p���n�N�������� Log �ɫ�R���A�H�K���P�޷| Key �ȭ��Ʒ�X��
				isql = "select * from ann_resp_exp where seq = " & pseq & " and seq1 = '" & pseq1 & "' "
				isql = isql & " and pay_times = '" & request("apay_times"&i) & "' "
				frs.Open isql,conn,1,1
				if not frs.EOF then
					'�s�W�~�O�P�� Log �� table:ann_resp_exp_log
					tkey_field = "seq;seq1;pay_times"
					tkey_value = pseq & ";" & pseq1 & ";" & request("apay_times"&i)
					call insert_log_table(conn,"U",prgid,"ann_resp_exp",tkey_field,tkey_value)
					'�R���~�O�P����
					isql = "delete from ann_resp_exp where seq = " & pseq & " and seq1 = '" & pseq1 & "' "
					isql = isql & " and pay_times = '" & request("apay_times"&i) & "' "
					'Response.Write "�R���~�O�P���� table:ann_resp_exp <br>" & fsql & "<br>"	
					showlog("[<u>Server_savestep.vbs</u>].insert_ann_exp.6="&isql)
					conn.execute isql
				end if
				frs.Close
			end if
		next
	end if
	'response.end
	
	if cstr(actrlnum) = "0" and yearctrl_flag = "Y" then
		'--�R���e,���J�� ann_exp_log		
		call insert_log_table(conn,"D",prgid,"ann_exp","seq;seq1",pseq&";"&pseq1)		
			
		isql = " delete from ann_exp where seq= '" & pseq & "' and seq1 = '" & pseq1 & "'"
		'Response.Write "�R�������ި��� table:ann_exp <br>" & isql & "<br>"			
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

'�P�~�O�ި�J�� ann_resp_exp
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
		'�s�W�~�O�P����		
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
		'Response.Write "�s�W�~�O�P�ި��� table:ann_resp_exp <br>" & fsql & "<br>"
		'response.end	
		'end if
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_resp_exp.1="&fsql)
		conn.execute fsql		
	
		'�R���~�O�ި��� tabke:ann_exp
		
		'--�R���e,���J�� ann_exp_log
		'if session("scode") = "admin" then
		'response.write frs("ann_sqlno") & "<br>"
		'response.end	
		'end if
		call insert_log_table(conn,"D",prgid,"ann_exp","ann_sqlno",frs("ann_sqlno"))
		
		fsql = "delete from ann_exp where ann_sqlno='" & frs("ann_sqlno") & "'"
		'if session("scode") = "admin" then
		'Response.Write "�R���~�O�ި��� tabke:ann_exp <br>" & fsql & "<br>"			
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

'�g�J�~�O�ɫ��ܶi��
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

	'Response.Write "�ק�~�O�ި��� table:ann_exp <br>" & fsql & "<br>"
	showlog("[<u>Server_savestep.vbs</u>].update_ann_step_grade="&fsql)
	conn.execute fsql		
	
end function
'�s�W�i�ת��ݮש��� step_expd
function insert_step_expd(prs_sqlno) 
	dim i
	for i = 1 to request("codenum")
		if request("rs_class"&i) <> empty and request("rs_code"&i) <> empty and request("act_code"&i) <> empty then
			fsql = "insert into step_expd (rs_sqlno,rs_type,rs_class,rs_code,act_code) values("
			fsql = fsql & prs_sqlno & ",'" & request("rs_type") & "','" & request("rs_class"&i) & "','" & request("rs_code"&i) & "','" & request("act_code"&i) & "')"
			'Response.Write "�s�W�i�ת��ݮש��� table:step_expd <br>" & fsql & "<br>"			
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
'�s�W�ި������ ctrl_exp
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
				'Response.Write "�s�W�ި������ table:ctrl_exp <br>" & fsql & "<br>"			
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
'�P�ި�J�� resp_exp
function insert_resp_exp(prsqlno)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'Ū���P�޸��
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'�s�W�ި��� table:resp_exp
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
			
			'Response.Write "�s�W�P�ި��� table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp.1="&usql)
			conn.execute usql
				
			'�R�������ި��� tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "�R�������ި��� tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp.2="&usql)
			conn.execute usql
			
			'�P�ި�J�ɮɡA�P�ɱN�]�ʦ^���ɥ��^�Ъ��]�ʵ�X���ݳB�z
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
		'Ū���P�޸��
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'�s�W�ި��� table:resp_exp
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
			
			'Response.Write "�s�W�P�ި��� table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp2.1="&usql)
			conn.execute usql
				
			'�R�������ި��� tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "�R�������ި��� tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp2.2="&usql)
			conn.execute usql
			
			'�P�ި�J�ɮɡA�P�ɱN�]�ʦ^���ɥ��^�Ъ��]�ʵ�X���ݳB�z
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
'�P�ި�J�� resp_exp
function insert_resp_exp3(prsqlno,presp_type)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'Ū���P�޸��
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'�s�W�ި��� table:resp_exp
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
			
			'Response.Write "�s�W�P�ި��� table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp3.1="&usql)
			conn.execute usql
				
			'�R�������ި��� tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "�R�������ި��� tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp3.2="&usql)
			conn.execute usql
			
			'�P�ި�J�ɮɡA�P�ɱN�]�ʦ^���ɥ��^�Ъ��]�ʵ�X���ݳB�z
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
'�P�ި�J�� resp_exp
function insert_resp_exp4(prsqlno,presp_type,presp_remark)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'Ū���P�޸��
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'�s�W�ި��� table:resp_exp
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
			
			'Response.Write "�s�W�P�ި��� table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp4.1="&usql)
			conn.execute usql
				
			'�R�������ި��� tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "�R�������ި��� tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp4.2="&usql)
			conn.execute usql
			
			'�P�ި�J�ɮɡA�P�ɱN�]�ʦ^���ɥ��^�Ъ��]�ʵ�X���ݳB�z
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
'�{�ǰe��T�{�ɾP��(�P�ޭ�]:I:�e��T�{)
function insert_resp_exp_brconf(prsqlno,presp_grade)
	dim i
	set rsf = server.CreateObject("ADODB.RECORDSET")
	ar = split(prsqlno,";")
	'response.write prsqlno & "<BR>"
	'response.end
	for i = 0 to ubound(ar) -1
		'Ū���P�޸��
		isql = "select * from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
		'response.write isql & "<BR>"
		'response.end
		rsf.open isql,conn,1,1
		if not rsf.EOF then
			'�s�W�ި��� table:resp_exp
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
			
			'Response.Write "�s�W�P�ި��� table:resp_exp <br>" & usql & "<br>"
			'response.end
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_brconf.1="&usql)
			conn.execute usql
				
			'�R�������ި��� tabke:ctrl_exp
			usql = "delete from ctrl_exp where ctrl_sqlno='" & ar(i) & "'"
			'Response.Write "�R�������ި��� tabke:ctrl_exp <br>" & usql & "<br>"			
			showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_brconf.2="&usql)
			conn.execute usql
			
			'�P�ި�J�ɮɡA�P�ɱN�]�ʦ^���ɥ��^�Ъ��]�ʵ�X���ݳB�z
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


'�P�ި�J�� resp_exp for ���ץ�
function insert_resp_exp_end(pseq,pseq1)
	set rsf = server.CreateObject("ADODB.RECORDSET")
	fsql = "select * from ctrl_exp where seq="& pseq &" and seq1='"& pseq1 &"'"
	fsql = fsql & " and ctrl_type<>'B8'"
	rsf.open fsql,conn,1,1
	while not rsf.eof
		'�s�W�ި��� table:resp_exp
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
		'Response.Write "�s�W�P�ި��� table:resp_exp <br>" & usql & "<br>"
		'response.end
		showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_end.1="&usql)
		conn.execute usql
				
		'�R�������ި��� tabke:ctrl_exp
		usql = "delete from ctrl_exp where ctrl_sqlno='" & rsf("ctrl_sqlno") & "'"
		'Response.Write "�R�������ި��� tabke:ctrl_exp <br>" & usql & "<br>"			
		showlog("[<u>Server_savestep.vbs</u>].insert_resp_exp_end.2="&usql)
		conn.execute usql
			
		'�P�ި�J�ɮɡA�P�ɱN�]�ʦ^���ɥ��^�Ъ��]�ʵ�X���ݳB�z
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

'�P�~�O�ި�J�� ann_resp_exp
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
		'�s�W�~�O�P����		
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
		
		'Response.Write "�s�W�~�O�P�ި��� table:ann_resp_exp <br>" & fsql & "<br>"
		showlog("[<u>Server_savestep.vbs</u>].insert_ann_resp_exp_end.1="&fsql)
		conn.execute fsql
	
		'�R���~�O�ި��� tabke:ann_exp
		
		'--�R���e,���J�� ann_exp_log
		call insert_log_table(conn,"D",prgid,"ann_exp","ann_sqlno",frs("ann_sqlno"))
		frs.movenext
	wend
	frs.close
	fsql = "delete from ann_exp"
	fsql = fsql & " where seq = " & pseq
	fsql = fsql & " and seq1 = '" & pseq1 & "' "
	'Response.Write "�R���~�O�ި��� tabke:ann_exp <br>" & fsql & "<br>"			
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
'�J���ײ�����
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
	'Response.Write "�s�W���ײ����� table:exp_end <br>" & fsql & "<br>"			
	'Response.End 
	showlog("[<u>Server_savestep.vbs</u>].insert_exp_end="&fsql)
	conn.execute fsql

end function
'���y���s�W�� exp_attach
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
	fsql = fsql & ",'���y���','" & tscan_name & "','" & tscan_name & "' "
	fsql = fsql & ",'A',"
	if pchk_status = "Y1" then
		fsql = fsql & "'Y',"
	else
		fsql = fsql & "'N',"
	end if
	fsql = fsql & "'N','N',getdate(),'" & session("scode") & "') "

	'Response.Write "�s�W���y����� table:exp_attach <br>" & fsql & "<br>"
	showlog("[<u>Server_savestep.vbs</u>].insert_exp_scan="&fsql)
	conn.execute fsql		
end function
%>
