<%
'新增附屬案性檔 attcasecode_exp
function insert_attcasecode_exp(patt_sqlno,psend_type) 
	dim i
	for i = 1 to request("codenum")
		if request("rs_class"&i) <> empty and request("rs_code"&i) <> empty _
		and request("act_code"&i) <> empty then
			fsql = "insert into attcasecode_exp (att_sqlno,send_type,"
			fsql = fsql & "rs_type,rs_class,rs_code,act_code) values("
			fsql = fsql & patt_sqlno & ",'"& psend_type &"','" & request("rs_type") & "',"
			fsql = fsql & "'" & request("rs_class"&i) & "','" & request("rs_code"&i) & "',"
			fsql = fsql & "'" & request("act_code"&i) & "')"
			'Response.Write "新增附屬案性檔 table:attcasecode_exp <br>" & fsql & "<br>"			
			'response.end
			conn.execute fsql
		end if
	next
end function
function delete_attcasecode_exp(patt_sqlno) 
	fsql = "delete from attcasecode_exp where att_sqlno = '" & patt_sqlno & "'"
	conn.execute fsql
end function
'刪除交辦檔
function delete_case_exp()
	dim fsql
	call insert_log_table(conn,"D",prgid,"case_exp","in_scode;in_no",in_scode&";"&in_no)
	fsql = "delete from case_exp"
	fsql = fsql & " where in_no = " & in_no &" and in_scode = " & in_scode
	'Response.Write "刪除交辦檔 table:case_exp <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'新增委辦案性檔
function insert_case_exp2(pin_scode,pin_no,pmustDel,pnum) 
	call insert_log_table(conn,"U",prgid,"case_exp2","in_scode;in_no",pin_scode&";"&pin_no)
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from case_exp2 where in_no='"& pin_no &"'"
		if pin_scode <> empty then
			fsql = fsql & " and in_scode='"& pin_scode &"'"
		end if
		conn.execute fsql
	end if
	
	for i=0 to pnum
		if i=0 then
			tarcase = request("arcase0")
			tact_code = request("act_code0")
		else
			tarcase = request("secarcase"&i)
			tact_code = request("secact_code"&i)
		end if
		if tarcase<>empty and tact_code<>empty then
			fsql = " insert into case_exp2(in_scode,in_no,seqno,seq,seq1,step_code,secact_code"
			fsql = fsql & ",secarcase_num,service,fee,mark)"
			fsql = fsql & " values('" & tin_scode1 & "','" & in_no & "','" & i & "',"
			if request("newold")="N" then
				fsql = fsql & "null,"
			else
				fsql = fsql & request("seq") & ","
			end if
			fsql = fsql & "'" & request("seq1") & "',"
			if i=0 then
				fsql = fsql & "'" & request("arcase0") & "',"
				fsql = fsql & "'" & request("act_code0") &"',"
				fsql = fsql & "1,"
			else
				fsql = fsql & "'" & request("secarcase"&i) & "',"
				fsql = fsql & "'" & request("secact_code"&i) &"',"
				fsql = fsql & chknumzero(request("secarcase_num"&i)) & ","
			end if
			fsql = fsql & chknumzero(request("service"&i)) & ","
			fsql = fsql & chknumzero(request("fees"&i)) & ",null)"
			'Response.Write "委辦案性檔 table:case_exp2 <br>" & fsql & "<br>"
			conn.Execute fsql
		end if
	next
end function
'刪除次委辦案性檔
function delete_case_exp2()
	dim fsql
	call insert_log_table(conn,"D",prgid,"case_exp2","in_no",in_no)
	fsql = "delete from case_exp2"
	fsql = fsql & " where in_no = " & in_no
	'Response.Write "刪除次委辦案性檔 table:case_exp2 <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'新增交辦申請人檔
function insert_exp_apcust(pin_scode,pin_no,pmustDel,pnum) 
	call insert_log_table(conn,"U",prgid,"exp_apcust","in_scode;in_no",pin_scode&";"&pin_no)
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from exp_apcust where in_no='"& pin_no &"'"
		'if pin_scode <> empty then
		'	fsql = fsql & " and in_scode='"& pin_scode &"'"
		'end if
		'Response.Write "交辦申請人檔 table:exp_apcust <br>" & fsql & "<br>"
		conn.execute fsql
	end if
	for i=1 to pnum
		if request("hapdel_flag"&i)="Y" then
			if trim(request("apcust_no"&i))<>empty then
				fsql = "update exp_apcust_log set ud_flag='D'"
				fsql = fsql & " where in_no='"& pin_no &"' and apsqlno='"& request("apsqlno"&i) &"'"
				'Response.Write fsql & "<br>"
				conn.execute fsql
			end if
		else
			fsql = " insert into exp_apcust(seq,seq1,in_scode,in_no,kind,apsqlno,ap_sql,"
			fsql = fsql & "ap_cname1,ap_cname2,ap_ename1,ap_ename2,title,brith,ap_crep,ap_erep,"
			fsql = fsql & "tran_scode,tran_date) values("
			if request("newold")<>"N" then
				fsql = fsql & seq & ",'" & seq1 & "'"
			else
				fsql = fsql & "null,'" & seq1 & "'"
			end if
			fsql = fsql & ",'" & pin_scode & "','" & pin_no & "'"
			fsql = fsql & ",'" & request("hapkind"&i) & "'," & request("apsqlno"&i) &","& chknumzero(request("ap_sql"&i))
			fsql = fsql & "," & chkcharnull2(request("ap_cname1_"&i)) & "," & chkcharnull2(request("ap_cname2_"&i)) 
			fsql = fsql & "," & chkcharnull(request("ap_ename1_"&i)) & "," & chkcharnull(request("ap_ename2_"&i))
			fsql = fsql & "," & chkcharnull(request("ap_title"&i)) & "," & chkdatenull(request("ap_brith"&i)) 
			fsql = fsql & "," & chkcharnull2(request("ap_crep"&i)) & "," & chkcharnull2(request("ap_erep"&i)) 
			fsql = fsql & ",'" & session("scode") & "',getdate())"
			'Response.Write "交辦申請人檔 table:exp_apcust <br>" & fsql & "<br>"
			'response.end
			
			conn.execute fsql
		end if
	next
	'response.end
end function
'刪除交辦申請人檔
function delete_exp_apcust()
	dim fsql
	call insert_log_table(conn,"D",prgid,"exp_apcust","in_no",in_no)
	fsql = "delete from exp_apcust"
	fsql = fsql & " where in_no = " & in_no
	'Response.Write "刪除交辦申請人檔 table:exp_apcust <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'新增交辦發明人檔
function insert_exp_ant(pin_scode,pin_no,pmustDel,pnum) 
	call insert_log_table(conn,"U",prgid,"exp_ant","in_scode;in_no",pin_scode&";"&pin_no)
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from exp_ant where in_no='"& pin_no &"'"
		'if pin_scode <> empty then
		'	fsql = fsql & " and in_scode='"& pin_scode &"'"
		'end if
		conn.execute fsql
	end if
	for i=1 to pnum
		if request("hantdel_flag"&i)="Y" then
			if trim(request("ant_no"&i)<>empty) then
				fsql = "update exp_ant_log set ud_flag='D'"
				fsql = fsql & " where in_no='"& pin_no &"' and ant_id='"& request("ant_no"&i) &"'"
				conn.execute fsql
			end if
		else
			fsql = "insert into exp_ant(seq,seq1,in_scode,in_no,kind,antsqlno,ant_id"
			fsql = fsql & ",ant_cname1,ant_cname2,antcomp,tran_date,tran_scode,ant_no)"
			fsql = fsql & " values("
			if request("newold")<>"N" or seq<>empty then
				fsql = fsql & seq & ",'" & seq1 & "'"
			else
				fsql = fsql & "null,'" & seq1 & "'"
			end if
			fsql = fsql & ",'" & pin_scode & "','" & pin_no & "','" & request("hantkind"&i) & "'"
			fsql = fsql & ","& chkcharnull(request("antsqlno"&i))
			fsql = fsql & ","& chkcharnull(request("ant_id"&i)) 
			fsql = fsql & ","& chkcharnull2(request("ant_cname1_"&i)) &","& chkcharnull2(request("ant_cname2_"&i))
			fsql = fsql & ","& chkcharnull(request("antcomp"&i)) 
			fsql = fsql & ","& "getdate(),'"& session("scode") &"',"& chkcharnull(request("ant_no"&i)) &")"
			'Response.Write "交辦發明人檔 table:exp_ant <br>" & fsql & "<br>"
			conn.execute fsql
		end if
	next
	'response.end
end function
'刪除交辦發明人檔
function delete_exp_ant()
	dim fsql
	call insert_log_table(conn,"D",prgid,"exp_ant","in_no",in_no)
	fsql = "delete from exp_ant"
	fsql = fsql & " where in_no = " & in_no
	'Response.Write "刪除交辦發明人檔 table:exp_ant <br>" & fsql & "<br>"
	conn.execute fsql
end function 
%>
