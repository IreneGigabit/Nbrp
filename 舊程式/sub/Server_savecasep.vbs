<%
'新增附屬案性檔
function insert_step_dmpd_temp(ptemp_step_sqlno,pnum,ptask_flag)
	dim i
	for i = 1 to pnum
		fsave_flag="N"
		if ptask_flag="N" then '可不輸入案性 brp2h
			fsave_flag="Y"
		else
			if request("rs_class"&i) <> empty and request("rs_code"&i) <> empty _
			and request("act_code"&i) <> empty then
				fsave_flag="Y"
			end if
		end if
		if fsave_flag="Y" then
			fsql = "insert into step_dmpd_temp (temp_step_sqlno,"
			fsql = fsql & "rs_type,rs_class,rs_code,act_code,tran_date,tran_scode) values("
			fsql = fsql & ptemp_step_sqlno & ",'" & request("rs_type") & "',"
			fsql = fsql & "'" & request("rs_class"&i) & "','" & request("rs_code"&i) & "',"
			fsql = fsql & "'" & request("act_code"&i) & "',getdate(),'"& session("scode") &"')"
			Response.Write "新增附屬案性檔 table:step_dmpd_temp <br>" & fsql & "<br>"			
			conn.execute fsql
		end if
	next
	'response.end
end function
function delete_step_dmpd_temp(ptemp_step_sqlno) 
	fsql = "delete from step_dmpd_temp where temp_step_sqlno = '" & ptemp_step_sqlno & "'"
	conn.execute fsql
end function
'新增附屬案性檔 attcasecode_dmp
function insert_attcasecode_dmp(patt_sqlno,psend_type) 
	dim i
	for i = 1 to request("codenum")
		if request("rs_class"&i) <> empty and request("rs_code"&i) <> empty _
		and request("act_code"&i) <> empty then
			fsql = "insert into attcasecode_dmp (att_sqlno,send_type,"
			fsql = fsql & "rs_type,rs_class,rs_code,act_code,tran_date,tran_scode) values("
			fsql = fsql & patt_sqlno & ",'"& psend_type &"','" & request("rs_type") & "',"
			fsql = fsql & "'" & request("rs_class"&i) & "','" & request("rs_code"&i) & "',"
			fsql = fsql & "'" & request("act_code"&i) & "',getdate(),'"& session("scode") &"')"
			'Response.Write "新增附屬案性檔 table:attcasecode_dmp <br>" & fsql & "<br>"			
			'response.end
			conn.execute fsql
		end if
	next
end function
function delete_attcasecode_dmp(patt_sqlno) 
	fsql = "delete from attcasecode_dmp where att_sqlno = '" & patt_sqlno & "'"
	conn.execute fsql
end function
'新增委辦案性檔
function insert_case_dmp2(pin_scode,pin_no,pmustDel,pnum) 
	call insert_log_table(conn,"U",prgid,"case_dmp2","in_scode;in_no",pin_scode&";"&pin_no)
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from case_dmp2 where in_no='"& pin_no &"'"
		if pin_scode <> empty then
			fsql = fsql & " and in_scode='"& pin_scode &"'"
		end if
		conn.execute fsql
	end if
	'response.write fsql & "<BR>"
	'response.write "pnum="&pnum & "<BR>"
	for i=0 to pnum
		if i=0 then
			tarcase = request("arcase")
			tact_code = request("act_code0")
		else
			tarcase = request("secarcase"&i)
			tact_code = request("secact_code"&i)
		end if
		'response.write "tarcase="& tarcase & "<BR>"
		'response.write "tact_code="& tact_code & "<BR>"
		if tarcase<>empty and tact_code<>empty then
			'response.write "i="& i & "<BR>"
			fsql = " insert into case_dmp2(in_scode,in_no,seqno,seq,seq1,step_code,secact_code"
			fsql = fsql & ",secarcase_num,service,fee,mark)"
			fsql = fsql & " values('" & pin_scode & "','" & in_no & "','" & i & "',"
			if request("newold")="N" then
				fsql = fsql & "null,"
			else
				fsql = fsql & request("seq") & ","
			end if
			fsql = fsql & "'" & request("seq1") & "',"
			if i=0 then
				fsql = fsql & "'" & request("arcase") & "',"
				fsql = fsql & "'" & request("act_code0") &"',"
				fsql = fsql & "1,"
				fsql = fsql & chknumzero(request("service")) & ","
				fsql = fsql & chknumzero(request("fees")) & ",null)"
			else
				fsql = fsql & "'" & request("secarcase"&i) & "',"
				fsql = fsql & "'" & request("secact_code"&i) &"',"
				fsql = fsql & chknumzero(request("secarcase_num"&i)) & ","
				fsql = fsql & chknumzero(request("service"&i)) & ","
				fsql = fsql & chknumzero(request("fees"&i)) & ",null)"
			end if
			'response.Write "委辦案性檔 table:case_dmp2 <br>" & fsql & "<br>"
			conn.Execute fsql
		end if
	next
	'response.end
end function
'刪除次委辦案性檔
function delete_case_dmp2()
	dim fsql
	call insert_log_table(conn,"D",prgid,"case_dmp2","in_no",in_no)
	fsql = "delete from case_dmp2"
	fsql = fsql & " where in_no = " & in_no
	'Response.Write "刪除次委辦案性檔 table:case_dmp2 <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'備註
function update_dmp_mark()
	dim i
	dim fsql
	dim rsf
	SET rsf = server.CreateObject("ADODB.RECORDSET")
	fsql = "select * from dmp_mark"
	if dmp_sqlno<>empty then
		fsql = fsql & " where dmp_sqlno = " & dmp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	if session("scode")="m983" then
	    'response.write fsql & "<BR>"
	    'response.end
	end if
	rsf.open fsql,conn,1,1
	if not rsf.eof then
		if dmp_sqlno<>empty then
			call insert_log_table(conn,"U",prgid,"dmp_mark","dmp_sqlno",dmp_sqlno)
		else
			call insert_log_table(conn,"U",prgid,"dmp_mark","seq;seq1",seq&";"&seq1)
		end if
		fsql = "update dmp_mark set family_master="& chkcharnull(request("family_master"))
		if trim(request("family_master"))="N" then
			fsql = fsql & ",family_dept="& chkcharnull(request("family_dept"))
			fsql = fsql & ",family_seq="& chknumzero(request("family_seq")) &",family_seq1="& chkcharnull(request("family_seq1"))
		else
			fsql = fsql & ",family_dept=''"
			fsql = fsql & ",family_seq=null,family_seq1=''"
		end if
		fsql = fsql & ",family_flag="& chkcharnull(request("family_flag"))
		if trim(request("family_flag"))="Y" then
			fsql = fsql & ",family_country="& chkcharnull(request("family_country")) &",family_apply_no="& chkcharnull(request("family_apply_no"))
		else
			fsql = fsql & ",family_country='',family_apply_no=''"
		end if
		fsql = fsql & ",tran_date=getdate(),tran_scode="& chkcharnull(session("scode"))
		if dmp_sqlno<>empty then
			fsql = fsql & " where dmp_sqlno = " & dmp_sqlno
		else
			fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
		end if
	else
		'舊案會有沒有dmp_mark的狀況，要insert一筆
        if request("seq")<>empty then
            if request("seq")<>0 then
                seq = request("seq")
                seq1 = request("seq1")
            end if
        end if		
		fsql = "insert into dmp_mark(dmp_sqlno,seq,seq1,family_master,family_dept,family_seq,family_seq1,"
		fsql = fsql & "family_flag,family_country,family_apply_no,"
		fsql = fsql & "tran_date,tran_scode)"
		fsql = fsql & " values("
		if request("dmp_sqlno")<>empty then
			fsql = fsql & chknumzero(dmp_sqlno) &",'"& seq &"','"& seq1 &"'"
		else
			if request("newold")="N" or request("newold")="S" then
				if request("newold")="N" then
					fsql = fsql & chknumzero(dmp_sqlno) &",null,''"
				else
					fsql = fsql & chknumzero(dmp_sqlno) &","& seq &",'"& seq1 &"'"
				end if
			else
				fsql = fsql & chknumzero(dmp_sqlno) &","& seq &",'"& seq1 &"'"
			end if
		end if
		fsql = fsql & ","& chkcharnull(request("family_master"))
		if trim(request("family_master"))="N" then
			fsql = fsql & ","& chkcharnull(request("family_dept"))
			fsql = fsql & ","& chknumzero(request("family_seq")) &","& chkcharnull(request("family_seq1"))
		else
			fsql = fsql & ",''"
			fsql = fsql & ",null,''"
		end if
		fsql = fsql & ","& chkcharnull(request("family_flag"))
		if trim(request("family_flag"))="Y" then
			fsql = fsql & ","& chkcharnull(request("family_country")) &","& chkcharnull(request("family_apply_no"))
		else
			fsql = fsql & ",'',''"
		end if
		fsql = fsql & ",getdate(),"& chkcharnull(session("scode"))
		fsql = fsql & ")"
	end if
	if session("scode")="m983" then
	    'Response.Write "更新案件備註檔 table:dmp_mark <br>" & fsql & "<br>"
	    'response.end 
	end if
	conn.execute fsql
end function
%>
