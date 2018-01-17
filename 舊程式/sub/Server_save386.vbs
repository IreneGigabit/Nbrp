<%'入386主檔
'-----exp, exp2, ap_exp
Set conni = Server.CreateObject("ADODB.Connection")
conni.Open session("sinbrp")
'response.write session("sinbrp")&"<BR>"
'response.end
'--新增案件主檔
function insert_exp_386(pseq,pseq1)
	dim fseq
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = " select * from exp where exp_sqlno = " & exp_sqlno
	'response.write fsql & "<BR>"
	rsf.Open fsql,conn,1,1
	fsql = "insert into exp(seq,seq1,cappl_name,eappl_name,cust_area,cust_seq,cust_seq1,att_sql," 
	fsql = fsql & "country,class,scode,end_code,tp_no,tp_no1,step_grade)"
	fsql = fsql & " values("& pseq &",'"& pseq1 &"',"
	fsql = fsql & chkcharnull3(rsf("cappl_name")) &","& chkcharnull3(rsf("eappl_name")) &","
	fsql = fsql & chkcharnull3(rsf("cust_area")) &","
	fsql = fsql & chknumzero(rsf("cust_seq")) &",0,"& chknumzero(rsf("att_sql")) &","
	fsql = fsql & chkcharnull3(rsf("country")) &","
	fsql = fsql & chkcharnull3(rsf("case1")) &","& chkcharnull3(rsf("scode1")) &","
	fsql = fsql & chkcharnull3(rsf("end_code")) &","
	fsql = fsql & chknumzero(rsf("tp_no")) &","& chkcharnull3(rsf("tp_no1")) &","
	fsql = fsql & chknumzero(request("nstep_grade")) & ")"
	'Response.Write "informix新增案件主檔 table:exp <br>" & fsql & "<br>"
	'Response.End 
	conni.execute fsql
	rsf.close
end function
'更新案件主檔
function update_exp_386()
	dim fseq
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = " select * from exp where seq = " & seq & " and seq1 = '" & seq1 & "'"
	rsf.Open fsql,conn,1,1
	fsql = "update exp set cappl_name=" & chkcharnull(rsf("cappl_name")) 
	fsql = fsql & ",eappl_name=" & chkcharnull(rsf("eappl_name"))
	fsql = fsql & ",cust_area="& chkcharnull2(rsf("cust_area")) 
	fsql = fsql & ",cust_seq=" & chknumzero(rsf("cust_seq")) &",cust_seq1=0,att_sql="& chknumzero(rsf("att_sql")) 
	fsql = fsql & ",country="& chkcharnull2(rsf("country")) &",class="& chkcharnull2(rsf("case1")) 
	fsql = fsql & ",scode=" & chkcharnull2(rsf("scode1"))
	fsql = fsql & ",end_code=" & chkdatenull(rsf("end_code"))
	fsql = fsql & ",tp_no=" & chknumzero(rsf("tp_no")) & ",tp_no1=" & chkcharnull2(rsf("tp_no1"))  
	fsql = fsql & ",step_grade="& chknumzero(request("nstep_grade"))
	fsql = fsql & " where seq="& seq &" and seq1='"& seq1 &"'"
	if session("scode")="m983" then
	    'Response.Write "informix更新案件檔 table:exp <br>" & fsql & "<br>"
	    'response.end
	end if
	conni.execute fsql
	rsf.close
end function
'案件申請人檔
function insert_ap_exp_386(pseq,pseq1)
	dim i
	dim fsql
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = " select a.*,b.apcust_no from exp_apcust a,apcust b"
	fsql = fsql & " where a.apsqlno=b.apsqlno and in_no = " & in_no
	'response.write fsql & "<BR>"
	'response.end 
	rsf.Open fsql,conn,1,1
	while not rsf.eof 
		fsql = "insert into ap_exp(sqlno,seq,seq1,apsqlno,apcust_no,ap_cname)"
		fsql = fsql & " values(0,"& pseq &",'"& pseq1 &"'"
		fsql = fsql & ","& chknumzero(rsf("apsqlno")) & ","& chkcharnull(rsf("apcust_no"))
		fsql = fsql & ","& chkcharnull(rsf("ap_cname1")&rsf("ap_cname2"))
		fsql = fsql & ")"
		'Response.Write "informix更新案件申請人檔 table:ap_exp <br>" & fsql & "<br>"
		'response.end
		conni.execute fsql
		rsf.movenext
	wend
	rsf.close
end function
'案件申請人檔
function insert_ap_exp_386_2(pseq,pseq1)
	dim i
	dim fsql
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = "select * From ap_exp where exp_sqlno = " & exp_sqlno
	'response.write fsql & "<BR>"
	'response.end 
	rsf.Open fsql,conn,1,1
	while not rsf.eof 
		fsql = "insert into ap_exp(sqlno,seq,seq1,apsqlno,apcust_no,ap_cname)"
		fsql = fsql & " values(0,"& pseq &",'"& pseq1 &"'"
		fsql = fsql & ","& chknumzero(rsf("apsqlno")) & ","& chkcharnull(rsf("apcust_no"))
		fsql = fsql & ","& chkcharnull(rsf("ap_cname1")&rsf("ap_cname2"))
		fsql = fsql & ")"
		'Response.Write "informix更新案件申請人檔 table:ap_exp <br>" & fsql & "<br>"
		'response.end
		conni.execute fsql
		rsf.movenext
	wend
	rsf.close
end function
'案件申請人檔
function update_ap_exp_386(pmustDel)
	dim i
	dim fsql
	Set rsf = Server.CreateObject("ADODB.Recordset")
	if pmustDel = "D" then
		fsql = "delete from ap_exp where seq='"& seq &"' and seq1='"& seq1 &"'"
		'response.write "informix: "& fsql & "<BR>"
		conni.execute fsql
	end if
	fsql = " select * from ap_exp where exp_sqlno = " & exp_sqlno
	'response.write fsql & "<BR>"
	rsf.Open fsql,conn,1,1
	while not rsf.eof 
		fsql = "insert into ap_exp(sqlno,seq,seq1,apsqlno,apcust_no,ap_cname)"
		fsql = fsql & " values(0,"& seq &",'"& seq1 &"'"
		fsql = fsql & ","& chknumzero(rsf("apsqlno")) & ","& chkcharnull(rsf("apcust_no"))
		fsql = fsql & ","& chkcharnull(rsf("ap_cname1")&rsf("ap_cname2"))
		fsql = fsql & ")"
		'Response.Write "informix更新案件申請人檔 table:ap_exp <br>" & fsql & "<br>"
		'response.end
		conni.execute fsql
		rsf.movenext
	wend
	rsf.close
end function
'請款資料寫入informix.account.eplus_temp
function insert_eplus_temp(pcase_no,pexch_no,pmustDel)
	dim fsql
	if pmustDel = "D" then
		fsql = "delete from account:eplus_temp where exch_no='" & pexch_no & "' and dept='P' "
		'response.write fsql & "<BR>"
		'response.end
		conni.Execute(fsql)
	end if
	
	'入區所帳款系統account.eplus_temp
	fail_flag=trim(request("hfees_stat"))
	if fail_flag=empty or isnull(fail_flag) then ext_flag="N"
	
	fsql = "insert into account:eplus_temp (dept,seq,seq1,step_grade,ext_seq,ext_seq1,"
	fsql = fsql & "case_no,country,arcase_type,arcase,exch_no,currency,db_money,hand_fee,pos_fee,"
	fsql = fsql & "fees,fail_flag,in_date,in_scode,chk_code) values "
	fsql = fsql & "('P'," & seq & ",'" & seq1 & "'," & request("nstep_grade") & ","
	fsql = fsql & request("tp_no") & ",'" & request("tp_no1") & "','" & pcase_no & "',"
	fsql = fsql & "'" & request("country") & "','" & request("rs_type") & "',"
	fsql = fsql & "'" & request("arcase"&i) & "','" & pexch_no & "','" & request("dn_currency") & "',"
	fsql = fsql & request("dn_money") & "," & request("hand_fee") & "," & request("pos_fee") & ","
	fsql = fsql & request("fees"&i) & ",'" & ext_flag & "'," & chkdatenull(date()) & ","
	fsql = fsql & "'" & session("se_scode") & "','N')"
	'Response.Write fsql&"<br>"
	'Response.End
	conni.Execute(fsql)
end function
%>
