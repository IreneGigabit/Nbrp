<%'入國外所資料庫
'session("sifedbs")
'-----入exp_temp
function insert_fexp_temp(prs_sqlno,pexp_sqlno,pseq,pseq1,patt_sqlno,ptp_no,ptp_no1,pcase_no)
	dim fsql
	set rsbr = server.CreateObject("ADODB.Recordset")
	set rsbr1 = server.CreateObject("ADODB.Recordset")

	fsql = "select * from exp where seq='"& pseq &"' and seq1='"& pseq1 &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	if not rsbr.EOF then
		fsql = "insert into exp_temp(seq,seq1,br_in_date,br_in_scode,br_rs_sqlno,"
		fsql = fsql & "case_no,att_sqlno,br_branch,br_no,br_no1,br_sales,br_eng,cust_area,cust_seq,"
		fsql = fsql & "att_sql,agent_no,agent_no1,country,cappl_name,eappl_name,jappl_name,"
		
		If Trim(pseq1)="T" Then
		    fsql = fsql & "your_no,"
		End If
		
		fsql = fsql & "in_date,case1,case_kind,ap_level,custprod_no,cust_prod,"
		fsql = fsql & "epctcountry,epctcoun_code,apply_date,apply_no,baseapply_date,cntry_date,pct_no," 
		fsql = fsql & "pub_date,pub_no,tapply_date,tapply_no,tbranch,tseq,tseq1,"
		fsql = fsql & "change_date,change_no,change_case1,appr_date,open_date,open_no,"
		fsql = fsql & "issue_date,issue_no,issue_no2,term1,term2,priority,"
		fsql = fsql & "gpay_flag,transfer_flag,ann_end_code,pay_times,pay_date,"
		fsql = fsql & "end_date,end_code,endremark,same_apply,same_first,same_br_no,same_br_no1,"
		fsql = fsql & "same_seq,same_seq1,auth_chk,pic_scode,pic_file_branch,pic_file1_path,"
		fsql = fsql & "pay_es,am_apply_date,am_ig,case_source,pr_scode,rent,"
		fsql = fsql & "announce_flag,announce_date,announce_cer,seu_issue_no,eeu_issue_no,"
		fsql = fsql & "entity_flag,inspect_flag,prior_no,pr_scodee)"
		fsql = fsql & " values("& chkcharnull(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & "getdate(),'"& session("scode") &"',"& prs_sqlno &","
		fsql = fsql & chkcharnull(pcase_no) &","& patt_sqlno &","
		fsql = fsql & chkcharnull(session("se_branch")) &","& pseq &",'"& pseq1 &"',"
		fsql = fsql & chkcharnull(rsbr("scode1")) &","& chkcharnull(rsbr("pr_scode")) &","
		fsql = fsql & chkcharnull(rsbr("cust_area")) &","
		fsql = fsql & chknumzero(rsbr("cust_seq")) &","& chknumzero(rsbr("att_sql")) &","
		fsql = fsql & chkcharnull(rsbr("expagt_no")) &","& chkcharnull(rsbr("expagt_no1")) &","
		fsql = fsql & chkcharnull(rsbr("country")) &","& chkcharnull2(rsbr("cappl_name")) &","
		fsql = fsql & chkcharnull2(rsbr("eappl_name")) &","& chkcharnull2(rsbr("jappl_name")) &","
		
		If Trim(pseq1)="T" Then
		    fsql = fsql & chkcharnull(rsbr("agent_rno")) & ","
		End If
				
		fsql = fsql & chkdatenull(rsbr("in_date")) &","& chkcharnull(rsbr("case1")) &","
		fsql = fsql & chkcharnull(rsbr("case_kind")) &","& chkcharnull(rsbr("ap_level")) &","
		fsql = fsql & chkcharnull(rsbr("custprod_no")) &","& chkcharnull(rsbr("cust_prod")) &"," 
		fsql = fsql & chkcharnull(rsbr("epctcountry")) &","& chkcharnull(rsbr("epctcoun_code")) &"," 
		fsql = fsql & chkdatenull(rsbr("apply_date")) &","& chkcharnull(rsbr("apply_no")) &","
		fsql = fsql & chkdatenull(rsbr("baseapply_date")) &","
		fsql = fsql & chkdatenull(rsbr("cntry_date")) &","& chkcharnull(rsbr("pct_no")) &","
		fsql = fsql & chkdatenull(rsbr("pub_date")) &","& chkcharnull(rsbr("pub_no")) &","
		fsql = fsql & chkdatenull(rsbr("tapply_date")) &","
		fsql = fsql & chkcharnull(rsbr("tapply_no")) &","& chkcharnull(rsbr("tbranch")) &","
		fsql = fsql & chknumzero(rsbr("tseq")) &","& chkcharnull(rsbr("tseq1")) &","
		fsql = fsql & chkdatenull(rsbr("change_date")) &","& chkcharnull(rsbr("change_no")) &","
		fsql = fsql & chkcharnull(rsbr("change_case1")) &","
		fsql = fsql & chkdatenull(rsbr("appr_date")) &","
		fsql = fsql & chkdatenull(rsbr("open_date")) &","& chkcharnull(rsbr("open_no")) &","
		fsql = fsql & chkdatenull(rsbr("issue_date")) &","& chkcharnull(rsbr("issue_no")) &","
		fsql = fsql & chkcharnull(rsbr("issue_no2")) &","
		fsql = fsql & chkdatenull(rsbr("term1")) &","& chkdatenull(rsbr("term2")) &","
		fsql = fsql & chkcharnull(rsbr("priority")) &","& chkcharnull(rsbr("gpay_flag")) &","
		fsql = fsql & chkcharnull(rsbr("transfer_flag")) &","& chkcharnull(rsbr("ann_end_code")) &","
		fsql = fsql & chkcharnull(rsbr("pay_times")) &","& chkdatenull(rsbr("pay_date")) &","
		fsql = fsql & chkdatenull(rsbr("end_date")) &","& chkcharnull(rsbr("end_code")) &","
		fsql = fsql & chkcharnull(rsbr("endremark")) &","
		fsql = fsql & chkcharnull(rsbr("same_apply")) &","& chkcharnull(rsbr("same_first")) &","
		fsql = fsql & chknumzero(rsbr("same_seq")) &","& chkcharnull(rsbr("same_seq1")) &","
		'柳月2016/10/5增加判斷為後案same_first=N且前案same_seq<>0，因台南SPE20898沒輸入前案造成抓取國外所案號錯誤
		if rsbr("same_apply") = "Y" and rsbr("same_first") = "N" and rsbr("same_seq")<>0 then
			isql = "select seq,seq1 from exp where br_branch='"& session("se_branch") &"'"
			isql = isql & " and br_no='"& rsbr("same_seq") & "' and br_no1='"& rsbr("same_seq1") &"'"
			rsbr1.open isql,connf,1,1
			if not rsbr1.eof then
				fsql = fsql & chknumzero(rsbr1("seq")) &","& chkcharnull(rsbr1("seq1")) &","
			else
				fsql = fsql & "'','',"
			end if
			rsbr1.close
		else
			fsql = fsql & "'','',"
		end if
		fsql = fsql & chkcharnull(rsbr("auth_chk")) &","& chkcharnull(rsbr("pic_scode")) &","
		fsql = fsql & chkcharnull(rsbr("pic_file_branch")) &","& chkcharnull(rsbr("pic_file1_path")) &","
		fsql = fsql & chkcharnull(rsbr("pay_es")) &","
		fsql = fsql & chkdatenull(rsbr("am_apply_date")) &","& chkcharnull(rsbr("am_ig")) &","
		fsql = fsql & chkcharnull(rsbr("case_source")) &","& chkcharnull(rsbr("pr_scode")) &","
		fsql = fsql & chkcharnull(rsbr("rent")) &","
		fsql = fsql & chkcharnull(rsbr("announce_flag")) &","& chkdatenull(rsbr("announce_date")) &","
		fsql = fsql & chkcharnull(rsbr("announce_cer")) &","
		fsql = fsql & chkcharnull(rsbr("seu_issue_no")) &","& chkcharnull(rsbr("eeu_issue_no")) &","
		fsql = fsql & chkcharnull(rsbr("entity_flag")) &","& chkcharnull(rsbr("inspect_flag")) &","
		fsql = fsql & chkcharnull(rsbr("prior_no")) &","& chkcharnull(rsbr("pr_scodee")) 
		fsql = fsql & ")"
		'Response.Write "================ fexp:exp <br>" & fsql & "<br>"
		'Response.End 
		connf.execute fsql
	end if
	set rsbr = nothing
end function
'-----入exp_mark_temp
function insert_fexp_mark_temp(prs_sqlno,pexp_sqlno,pseq,pseq1,ptp_no,ptp_no1)
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select * from exp_mark where seq='"& pseq &"' and seq1='"& pseq1 &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	if not rsbr.EOF then
		fsql = "insert into exp_mark_temp(br_branch,br_rs_sqlno,seq,seq1,slang,spages,words"
		fsql = fsql & ",claim,draw_cnt,rep_draw,ipc,draw_paper,ind_item"
		fsql = fsql & ",copy_flag,copy_seq,copy_seq1,family_master,family_dept,family_seq,family_seq1"
		fsql = fsql & ",family_flag,family_country,family_apply_no,hk_flag,hk_seq,hk_seq1" 
		fsql = fsql & ",eu_seq,eu_seq1,fremark,bremark)"
		fsql = fsql & " values("& chkcharnull(session("se_branch")) &","& prs_sqlno &","
		fsql = fsql & chkcharnull(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & chkcharnull(rsbr("slang")) &","& chknumzero(rsbr("spages")) &","
		fsql = fsql & chknumzero(rsbr("words")) &","& chknumzero(rsbr("claim")) &","
		fsql = fsql & chknumzero(rsbr("draw_cnt")) &","& chkcharnull(rsbr("rep_draw")) &","
		fsql = fsql & chkcharnull(rsbr("ipc")) &","& chknumzero(rsbr("draw_paper")) &","
		fsql = fsql & chknumzero(rsbr("ind_item")) &","
		fsql = fsql & chkcharnull(rsbr("copy_flag")) &","& chknumzero(rsbr("copy_seq")) &","
		fsql = fsql & chkcharnull(rsbr("copy_seq1")) &","
		fsql = fsql & chkcharnull(rsbr("family_master")) &","& chkcharnull(rsbr("family_dept")) &","
		fsql = fsql & chknumzero(rsbr("family_seq")) &","& chkcharnull(rsbr("family_seq1")) &","
		fsql = fsql & chkcharnull(rsbr("family_flag")) &","& chkcharnull(rsbr("family_country")) &","
		fsql = fsql & chkcharnull(rsbr("family_apply_no")) &","& chkcharnull(rsbr("hk_flag")) &","
		fsql = fsql & chknumzero(rsbr("hk_seq")) &","& chkcharnull(rsbr("hk_seq1")) &","
		fsql = fsql & chknumzero(request("eu_seq")) & ","& chkcharnull(request("eu_seq1"))
		fsql = fsql & ","& chkcharnull2(rsbr("fremark")) &","& chkcharnull2(rsbr("fremark")) &")"
		'Response.Write "========== fexp:exp_mark <br>" & fsql & "<br>"
		'Response.End 
		connf.execute fsql
	end if
	set rsbr = nothing
end function
'-----入exp_prior_temp
function insert_fexp_prior_temp(prs_sqlno,pexp_sqlno,pseq,pseq1,ptp_no,ptp_no1)
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select * from exp2 where exp_sqlno='"& pexp_sqlno &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into exp_prior_temp(br_branch,br_rs_sqlno,seq,seq1"
		fsql = fsql & ",prior_date,prior_no,prior_country,prior_flag,apply_yn"
		fsql = fsql & ",apply_seq,apply_seq1,tran_date,tran_scode,prior_no2"
		fsql = fsql & ")"
		fsql = fsql & " values("& chkcharnull(session("se_branch")) &","& prs_sqlno &","
		fsql = fsql & chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & chkdatenull(rsbr("apply_date")) &","& chkcharnull(rsbr("apply_no")) &","
		fsql = fsql & chkcharnull(rsbr("country")) &","& chkcharnull(rsbr("prior_yn")) &","
		fsql = fsql & chkcharnull(rsbr("apply_yn")) &","
		fsql = fsql & chkcharnull(rsbr("apply_seq")) &"," & chkcharnull(rsbr("apply_seq1")) &","
		fsql = fsql & chkdatenull2(rsbr("tran_date")) & ","
		'fsql = fsql & "'"& FormatDateTime(rsbr("tran_date"),2) &" "& string(2-len(hour(rsbr("tran_date"))),"0") & hour(rsbr("tran_date")) &":"& string(2-len(minute(rsbr("tran_date"))),"0") & minute(rsbr("tran_date")) &":"& string(2-len(second(rsbr("tran_date"))),"0") & second(rsbr("tran_date")) &"',"
		fsql = fsql & chkcharnull(rsbr("tran_scode")) & ","& chkcharnull(rsbr("prior_no"))
		fsql = fsql & ")"
		'Response.Write "========== fexp:exp_prior_temp <br>" & fsql & "<br>"
		'Response.End 
		connf.execute fsql
		rsbr.movenext
	wend
	set rsbr = nothing
end function
'-----入exp_tec_temp
function insert_fexp_tec_temp(prs_sqlno,pexp_sqlno,pseq,pseq1,patt_sqlno,ptp_no,ptp_no1,pcase_no)
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	'Response.Write "========== fexp:exp_tec_temp <br>"
	fsql = "select * from exp_tec where seq='"& pseq &"' and seq1='"& pseq1 &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into exp_tec_temp(br_branch,br_rs_sqlno,br_case_no,att_sqlno,seq,seq1"
		fsql = fsql & ",tec_flag,tec_code,tec_seq,tec_seq1,tec_country,tec_apply_no"
		fsql = fsql & ")"
		fsql = fsql & " values("& chkcharnull(session("se_branch")) &","& prs_sqlno &","
		fsql = fsql & chkcharnull(pcase_no) &",'"& patt_sqlno &"',"
		fsql = fsql & chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & chkcharnull(rsbr("tec_flag")) &","& chkcharnull(rsbr("tec_code")) &","
		fsql = fsql & chknumzero(rsbr("tec_seq")) &","& chkcharnull(rsbr("tec_seq1")) &","
		fsql = fsql & chkcharnull(rsbr("tec_country")) &","& chkcharnull(rsbr("tec_apply_no")) 
		fsql = fsql & ")"
		'if session("scode")="admin" then
		'Response.Write fsql & "<br>"
		'Response.End 
		'end if
		connf.execute fsql
		rsbr.movenext
	wend
	set rsbr = nothing
end function
'-----入exp_apcust_temp
function insert_fexp_apcust_temp(prs_sqlno,pexp_sqlno,pseq,pseq1,pstep_grade,patt_sqlno,ptp_no,ptp_no1,pcase_no)
	set rsbr = server.CreateObject("ADODB.Recordset")
	set rsbr2 = server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select a.*,b.apclass,b.ap_country,b.ap_ename1,b.ap_ename2,b.ap_zip,b.ap_addr1,b.ap_addr2,"
	fsql = fsql & "b.ap_eaddr1,b.ap_eaddr2,b.ap_eaddr3,b.ap_eaddr4"
	fsql = fsql & " from ap_exp a,apcust b where a.seq='"& pseq &"' and a.seq1='"& pseq1 &"'"
	fsql = fsql & " and a.apcust_no=b.apcust_no"
	fsql = fsql & " order by sqlno"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into exp_apcust_temp(br_branch,br_rs_sqlno,br_case_no,att_sqlno,seq,seq1"
		fsql = fsql & ",apsqlno,apcust_no,apclass,ap_country,ap_cname1,ap_cname2,ap_sql"
		fsql = fsql & ",ap_ename1,ap_ename2,ap_zip,ap_addr1,ap_addr2,ap_eaddr1,ap_eaddr2,ap_eaddr3,ap_eaddr4"
		fsql = fsql & ",ap_crep,ap_erep,title,brith,tran_date,tran_scode"
		fsql = fsql & ")"
		fsql = fsql & " values("& chkcharnull(session("se_branch")) &","& prs_sqlno &","
		fsql = fsql & chkcharnull(pcase_no) &","& patt_sqlno &","
		fsql = fsql & chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & chknumzero(rsbr("apsqlno")) &","& chkcharnull(rsbr("apcust_no")) &","
		fsql = fsql & chkcharnull(rsbr("apclass")) &","& chkcharnull(rsbr("ap_country")) &","
		fsql = fsql & chkcharnull(rsbr("ap_cname1")) &","& chkcharnull(rsbr("ap_cname2")) &","
		fsql = fsql & chknumzero(rsbr("ap_sql")) &","
		if rsbr("ap_sql")=0 then
			fsql = fsql & chkcharnull(rsbr("ap_ename1")) &","& chkcharnull(rsbr("ap_ename2")) &","
			fsql = fsql & chkcharnull(rsbr("ap_zip")) &","
			fsql = fsql & chkcharnull(rsbr("ap_addr1")) &","& chkcharnull(rsbr("ap_addr2")) &","
			fsql = fsql & chkcharnull(rsbr("ap_eaddr1")) &","& chkcharnull(rsbr("ap_eaddr2")) &","
			fsql = fsql & chkcharnull(rsbr("ap_eaddr3")) &","& chkcharnull(rsbr("ap_eaddr4")) &","
		else
			fsqlq = "select * from ap_nameaddr where apsqlno="& rsbr("apsqlno")
			fsqlq = fsqlq & " and ap_sql="& rsbr("ap_sql")
			rsbr2.open fsqlq,conn,1,1
			if not rsbr2.eof then
				fsql = fsql & chkcharnull(rsbr2("ap_ename1")) &","& chkcharnull(rsbr2("ap_ename2")) &","
				fsql = fsql & chkcharnull(rsbr2("ap_zip")) &","
				fsql = fsql & chkcharnull(rsbr2("ap_addr1")) &","& chkcharnull(rsbr2("ap_addr2")) &","
				fsql = fsql & chkcharnull(rsbr2("ap_eaddr1")) &","& chkcharnull(rsbr2("ap_eaddr2")) &","
				fsql = fsql & chkcharnull(rsbr2("ap_eaddr3")) &","& chkcharnull(rsbr2("ap_eaddr4")) &","
			else
				fsql = fsql & chkcharnull(rsbr("ap_ename1")) &","& chkcharnull(rsbr("ap_ename2")) &","
				fsql = fsql & chkcharnull(rsbr("ap_zip")) &","
				fsql = fsql & chkcharnull(rsbr("ap_addr1")) &","& chkcharnull(rsbr("ap_addr2")) &","
				fsql = fsql & chkcharnull(rsbr("ap_eaddr1")) &","& chkcharnull(rsbr("ap_eaddr2")) &","
				fsql = fsql & chkcharnull(rsbr("ap_eaddr3")) &","& chkcharnull(rsbr("ap_eaddr4")) &","
			end if
			rsbr2.close
		end if		
		fsql = fsql & chkcharnull(rsbr("ap_crep")) &","& chkcharnull(rsbr("ap_erep")) &","
		fsql = fsql & chkcharnull(rsbr("ap_title")) &","& chkdatenull(rsbr("ap_brith")) &","
		if rsbr("tran_date")<>empty then
			fsql = fsql & "'"& FormatDateTime(rsbr("tran_date"),2) &" "& string(2-len(hour(rsbr("tran_date"))),"0") & hour(rsbr("tran_date")) &":"& string(2-len(minute(rsbr("tran_date"))),"0") & minute(rsbr("tran_date")) &":"& string(2-len(second(rsbr("tran_date"))),"0") & second(rsbr("tran_date")) &"',"
		else
			fsql = fsql & "null,"
		end if
		fsql = fsql & chkcharnull(rsbr("tran_scode"))
		fsql = fsql & ")"
		'Response.Write "========== fexp:exp_apcust_temp <br>" & fsql & "<br>"
		'Response.End 
		connf.execute fsql
		rsbr.movenext
	wend
	set rsbr = nothing
end function
'-----入exp_ant_temp
function insert_fexp_ant_temp(prs_sqlno,pexp_sqlno,pseq,pseq1,pstep_grade,patt_sqlno,ptp_no,ptp_no1,pcase_no)
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select a.*,b.apclass,b.ant_country,b.ant_ename1,b.ant_ename2,b.ant_zip,b.ant_addr1,b.ant_addr2,"
	fsql = fsql & "b.ant_eaddr1,b.ant_eaddr2,b.ant_eaddr3,b.ant_eaddr4"
	fsql = fsql & " from ap_expant a,inventor b where a.seq='"& pseq &"' and a.seq1='"& pseq1 &"'"
	fsql = fsql & " and a.ant_no=b.ant_no "
	fsql = fsql & " order by a.expant_sqlno"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into exp_ant_temp(br_branch,br_rs_sqlno,br_case_no,att_sqlno,seq,seq1"
		fsql = fsql & ",bantsqlno,antcomp,tran_date,tran_scode"
		fsql = fsql & ")"
		fsql = fsql & " values("& chkcharnull(session("se_branch")) &","& prs_sqlno &","
		fsql = fsql & chkcharnull(pcase_no) &","& patt_sqlno &","
		fsql = fsql & chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & chknumzero(rsbr("antsqlno")) &","& chkcharnull(rsbr("antcomp")) &","
		fsql = fsql & "'"& FormatDateTime(rsbr("tran_date"),2) &" "& string(2-len(hour(rsbr("tran_date"))),"0") & hour(rsbr("tran_date")) &":"& string(2-len(minute(rsbr("tran_date"))),"0") & minute(rsbr("tran_date")) &":"& string(2-len(second(rsbr("tran_date"))),"0") & second(rsbr("tran_date")) &"',"
		fsql = fsql & chkcharnull(rsbr("tran_scode"))
		fsql = fsql & ")"
		'Response.Write "========== fexp:exp_ant_temp <br>" & fsql & "<br>"
		'Response.End 
		connf.execute fsql
		rsbr.movenext
	wend
	set rsbr = nothing
end function
'-----入bexp_attach
function insert_fbattach_exp(prs_sqlno,pexp_sqlno,pseq,pseq1,pstep_grade,patt_sqlno,ptp_no,ptp_no1)
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	
	fsql = "select * from exp_attach where seq='"& pseq &"' and seq1='"& pseq1 &"' and attach_flag<>'D'"
	fsql = fsql & " and att_sqlno="& patt_sqlno
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into exp_attach_temp(seq,seq1,br_branch,br_in_date,br_in_scode"
		fsql = fsql & ",br_no,br_no1,br_step_grade,br_rs_sqlno,br_attach_sqlno"
		fsql = fsql & ",source,in_date,in_scode,attach_no,attach_path,doc_type"
		fsql = fsql & ",attach_desc,attach_name,source_name,exp_flag,attach_size,open_flag"
		fsql = fsql & ",mark,tran_date,tran_scode"
		fsql = fsql & ") values("
		fsql = fsql & chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","
		fsql = fsql & chkcharnull(session("se_branch")) &","
		fsql = fsql & chkdatenull2(rsbr("in_date")) &","& chkcharnull(rsbr("in_scode")) &","
		fsql = fsql & chknumzero(pseq) &","& chkcharnull(pseq1) &","& chknumzero(pstep_grade) &","
		fsql = fsql & prs_sqlno &"," & chknumzero(rsbr("attach_sqlno")) &",'Bsend',"
		fsql = fsql & "getdate(),'"& session("scode") &"',"
		fsql = fsql & chkcharnull(rsbr("attach_no")) &","& chkcharnull(rsbr("attach_path")) &","
		fsql = fsql & chkcharnull(rsbr("doc_type")) &","
		fsql = fsql & chkcharnull(rsbr("attach_desc")) &","& chkcharnull(rsbr("attach_name")) &","
		fsql = fsql & chkcharnull(rsbr("source_name")) &",'Y',"& chkcharnull(rsbr("attach_size")) &","
		fsql = fsql & chkcharnull(rsbr("mark")) &","& chkcharnull(rsbr("open_flag")) &","
		fsql = fsql & "'"& FormatDateTime(rsbr("tran_date"),2) &" "& string(2-len(hour(rsbr("tran_date"))),"0") & hour(rsbr("tran_date")) &":"& string(2-len(minute(rsbr("tran_date"))),"0") & minute(rsbr("tran_date")) &":"& string(2-len(second(rsbr("tran_date"))),"0") & second(rsbr("tran_date")) &"',"
		fsql = fsql & chkcharnull(rsbr("tran_scode"))
		fsql = fsql & ")"
		if session("scode")="n319" then
		'	Response.Write "========== fexp:exp_attach_temp <br>" & fsql & "<br>"
		'	response.end
		end if
		connf.execute fsql
		rsbr.movenext
	wend
	set rsbr = nothing
	'Response.End 
end function
'-----入step_exp_temp
function insert_fstep_exp_temp(prs_sqlno,pseq,pseq1,pstep_grade,patt_sqlno,pwork_opt) 
	set rsbr = server.CreateObject("ADODB.Recordset")
	set rsbr1 = server.CreateObject("ADODB.Recordset")
	set rsbr2 = server.CreateObject("ADODB.Recordset")
	dim fsql
	dim fsql1
	fsql = "select * from step_exp where seq='"& pseq &"' and seq1='"& pseq1 &"'"
	fsql = fsql & " and step_grade='" & pstep_grade &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	if not rsbr.EOF then
		isql = "select * from exp"
		isql = isql & " where seq='"& pseq &"' and seq1='"& pseq1 &"'"
		rsbr1.Open isql,conn,1,1
		atp_no = trim(rsbr1("tp_no"))
		rsbr1.Close 
		
		fsql = "insert into step_exp_temp(seq,seq1,al,rs,br_branch,br_in_date,br_rs_sqlno,"
		fsql = fsql & "case_no,att_sqlno,br_step_grade,br_step_date,br_in_scode,br_in_no,newold_type,"
		fsql = fsql & "case_type,rs_type,rs_class,rs_code,act_code,case_stat,rs_detail,doc_detail,"
		fsql = fsql & "pr_scode,br_main_rs_sqlno,tot_num,brconf_flag,brconf_sqlno,slang,remark,"
		fsql = fsql & "br_apply_date,br_ctrl_date,tran_date,tran_scode,rsagent_no,rsagent_no1"
		fsql = fsql & ",spay_times,epay_times,br_back_flag,br_end_flag,end_code,br_reason"
		fsql = fsql & ",urg_date,expagt_no,expagt_no1,expagt_reason,ctrl_flag,ctrl_case_no"
		fsql = fsql & ",ar_mark) values("
		fsql = fsql & chknumzero(tp_no) &","& chkcharnull(tp_no1) &","
		fsql = fsql & "'L','R',"
		fsql = fsql & chkcharnull(session("se_branch")) &",getdate(),"
		fsql = fsql & rsbr("rs_sqlno") &",'"& case_no &"',"& chknumzero(rsbr("att_sqlno")) &","
		fsql = fsql & nstep_grade &",'"& rsbr("step_date") &"',"
		fsql = fsql & "'"& in_scode &"','"& in_no &"',"
		'newold_type:A1新立案,A2新案指定編號,B後續案
		if left(pwork_opt,1) = "P" then 'PF:需回覆國外所「需收費之後續接洽/交辦」,P:需回覆國外所「不需收費之後續接洽/交辦」
			fsql = fsql & "'B1',"
			ctrlt_date = ""
			last_date = ""
		elseif pwork_opt = "SA" then 'SA:營洽新增交辦發文
			if atp_no<>empty and atp_no <>"0" then
				fsql = fsql & "'B1',"
			else
				fsql = fsql & "'A1',"
			end if
			ctrlt_date = ""
			last_date = ""
		else 'S:營洽交辦發文
			fsql1 = "select ctrlt_date,last_date,case_stat"
			fsql1 = fsql1 & " from case_exp where case_no='"& case_no &"'"
			rsbr1.open fsql1,conn,1,1
			if not rsbr1.eof then
				if rsbr1("case_stat")="N" then
					'曾聯發過且已有國外所編號視為舊案
					fsql1 = "select count(*) as cnt from step_exp"
					fsql1 = fsql1 & " where seq="& pseq &" and seq1='"& pseq1 &"'"					fsql1 = fsql1 & " and cg='T' and rs='S'"
					rsbr2.open fsql1,conn,1,1
					if not rsbr2.eof then
						if rsbr2("cnt")>0 then
							if atp_no<>empty and atp_no <>"0" then
								fsql = fsql & "'B1',"
							else
								fsql = fsql & "'A1',"
							end if
						else
							fsql = fsql & "'A1',"
						end if
					else
						fsql = fsql & "'A1',"
					end if
					rsbr2.close
				elseif rsbr1("case_stat")="S" then
					fsql = fsql & "'A2',"
				else
					fsql = fsql & "'B1',"
				end if
				ctrlt_date = rsbr1("ctrlt_date")
				last_date = rsbr1("last_date")
			else
				fsql = fsql & "'B1',"
				ctrlt_date = ""
				last_date = ""
			end if
			rsbr1.close
		end if
		fsql = fsql & chkcharnull(rsbr("send_way")) &","& chkcharnull(rsbr("rs_type")) &","
		fsql = fsql & chkcharnull(rsbr("rs_class")) &","& chkcharnull(rsbr("rs_code")) &","
		fsql = fsql & chkcharnull(rsbr("act_code")) &","& chkcharnull(rsbr("case_stat")) &","
		fsql = fsql & chkcharnull(rsbr("rs_detail")) &","& chkcharnull(rsbr("doc_detail")) &","
		fsql = fsql & chkcharnull(rsbr("pr_scode")) &","
		if left(pwork_opt,1) = "P" then
			fsql = fsql & prs_sqlno &",1,"
		else
			fsql = fsql & "0,1," '需另外update
		end if
		fsql1 = "select slang,brconf_flag,brconf_sqlno,remark"
		fsql1 = fsql1 & " from attcase_exp where att_sqlno='"& patt_sqlno &"'"
		rsbr1.open fsql1,conn,1,1
		if not rsbr1.eof then
			brconf_flag = rsbr1("brconf_flag")
			brconf_sqlno = rsbr1("brconf_sqlno")
			slang = rsbr1("slang")
			remark = rsbr1("remark")
		else
			brconf_flag = ""
			brconf_sqlno = 0
			slang = ""
			remark = ""
		end if
		rsbr1.close
		fsql = fsql & chkcharnull(brconf_flag) &","& chknumzero(brconf_sqlno) &","
		fsql = fsql & chkcharnull(slang) &","& chkcharnull2(remark) &","
		fsql = fsql & chkdatenull(ctrlt_date) &","& chkdatenull(last_date) &","
		'fsql = fsql & chkdatenull(rsbr("tran_date")) &","
		fsql = fsql & "'"& FormatDateTime(rsbr("tran_date"),2) &" "& string(2-len(hour(rsbr("tran_date"))),"0") & hour(rsbr("tran_date")) &":"& string(2-len(minute(rsbr("tran_date"))),"0") & minute(rsbr("tran_date")) &":"& string(2-len(second(rsbr("tran_date"))),"0") & second(rsbr("tran_date")) &"',"
		fsql = fsql & chkcharnull(rsbr("tran_scode"))
		fsql1 = "select expagt_no,expagt_no1"
		fsql1 = fsql1 & " from exp where seq='"& rsbr("seq") &"' and seq1='"& rsbr("seq1") &"'"
		'response.write fsql1
		'response.end
		rsbr1.open fsql1,conn,1,1
		if not rsbr1.eof then
			fsql = fsql & ","& chkcharnull(rsbr1("expagt_no")) &","& chkcharnull(rsbr1("expagt_no1")) &","
		else
			fsql = fsql & ",'','',"
		end if
		rsbr1.close	
		fsql = fsql & chkcharnull(rsbr("spay_times")) &","& chkcharnull(rsbr("epay_times"))
		fsql = fsql &","& chkcharnull(rsbr("back_flag")) &","& chkcharnull(rsbr("end_flag"))
		fsql1 = "select endremark from attcase_exp where att_sqlno='"& patt_sqlno &"'"
		rsbr1.open fsql1,conn,1,1
		if not rsbr1.eof then
			step_endremark = rsbr1("endremark")
		else
			step_endremark = ""
		end if
		rsbr1.close
		fsql = fsql &","& chkcharnull(step_endremark) &","& chkcharnull(rsbr("flag_remark"))
		fsql = fsql &","& chkdatenull(rsbr("urg_date"))
		fsql1 = "select expagt_no,expagt_no1,expagt_reason,ctrl_flag,ctrl_case_no,ar_mark"
		fsql1 = fsql1 & " from case_exp where case_no='"& case_no &"'"
		rsbr1.open fsql1,conn,1,1
		if not rsbr1.eof then
			fsql = fsql & ","& chkcharnull(rsbr1("expagt_no")) & ","& chkcharnull(rsbr1("expagt_no1"))
			fsql = fsql & ","& chkcharnull(rsbr1("expagt_reason"))
			fsql = fsql & ","& chkcharnull(rsbr1("ctrl_flag")) & ","& chkcharnull(rsbr1("ctrl_case_no"))
			fsql = fsql & ",'"& rsbr1("ar_mark") & "'"
		else
			fsql = fsql & ",'','','','','',''"
		end if
		rsbr1.close
		fsql = fsql & ")"
		'Response.Write "========== fexp:step_exp_temp <br>" & fsql & "<br>"
		'Response.End 
		connf.execute fsql
	end if
	set rsbr = nothing
end function
'-----入step_expd_temp
function insert_fstep_expd_temp(prs_sqlno,patt_sqlno,pcase_no) 
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select * from step_expd where rs_sqlno='"& prs_sqlno &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into step_expd_temp (br_branch,br_rs_sqlno,br_case_no,att_sqlno,"
		fsql = fsql & "rs_type,rs_class,rs_code,act_code) values("
		fsql = fsql & chkcharnull(session("se_branch")) &","& prs_sqlno &",'"& pcase_no &"',"
		fsql = fsql & patt_sqlno &",'" & rsbr("rs_type") & "',"
		fsql = fsql & "'" & rsbr("rs_class") & "','" & rsbr("rs_code") & "','" & rsbr("act_code") & "'"
		fsql = fsql & ")"
		'Response.Write "========== fexp:step_expd_temp <br>" & fsql & "<br>"
		connf.execute fsql
		rsbr.movenext
	wend
	'Response.End 
end function
'-----入step_ctrl_temp
function insert_fctrl_exp_temp(prs_sqlno,pseq,pseq1,pstep_grade,patt_sqlno,ptp_no,ptp_no1)	
	set rsbr = server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select * from ctrl_exp where seq='"& pseq &"' and seq1='"& pseq1 &"'"
	fsql = fsql & " and step_grade='" & pstep_grade &"'"
	'response.write fsql & "<BR>"
	'response.end
	rsbr.Open fsql,conn,1,1
	while not rsbr.EOF 
		fsql = "insert into ctrl_exp_temp (br_branch,br_rs_sqlno,seq,seq1,step_grade," 
		fsql = fsql & "ctrl_type,ctrl_remark,ctrl_date,date_ctrl,tran_date,tran_scode"
		fsql = fsql & ") values("
		fsql = fsql & "'"& session("se_branch") &"'," & prs_sqlno &","
		fsql = fsql & chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","& pstep_grade & ","
		fsql = fsql & "'" & rsbr("ctrl_type") & "'," & chkcharnull(rsbr("ctrl_remark")) & ","
		fsql = fsql & chkdatenull(rsbr("ctrl_date")) & "," & chkcharnull(rsbr("date_ctrl")) & ","
		fsql = fsql & "'"& FormatDateTime(rsbr("tran_date"),2) &" "& string(2-len(hour(rsbr("tran_date"))),"0") & hour(rsbr("tran_date")) &":"& string(2-len(minute(rsbr("tran_date"))),"0") & minute(rsbr("tran_date")) &":"& string(2-len(second(rsbr("tran_date"))),"0") & second(rsbr("tran_date")) &"',"
		fsql = fsql & chkcharnull(rsbr("tran_scode"))
		fsql = fsql & ")"
		'Response.Write "========== fexp:ctrl_exp_temp <br>" & fsql & "<br>"
		connf.execute fsql
		rsbr.movenext
	wend
	'Response.End 
end function
'-----入todo_exp
function insert_ftodo_exp(prs_sqlno,ptp_no,ptp_no1,pdowhat,ptodosqlno,pwork_opt)
	dim fsql
	Set rsbr1 = Server.CreateObject("ADODB.Recordset")
	Set rsbr2 = Server.CreateObject("ADODB.Recordset")
	fsql = "insert into todo_exp(syscode,apcode,br_branch,br_rs_sqlno,seq,seq1,"
	fsql = fsql & "in_date,in_scode,dowhat,job_status,br_apcode,br_todosqlno)"
	fsql = fsql & " values('FEXP','"& prgid &"','"& session("se_branch") &"',"
	fsql = fsql & chknumzero(prs_sqlno) &","& chknumzero(ptp_no) &","& chkcharnull(ptp_no1) &","
	fsql = fsql & "getdate(),'"& session("scode") &"',"
	if left(pwork_opt,1) = "P" then 'newold_type:A1新立案,A2新案指定編號,B後續案
		fsql = fsql & "'LR_O',"
	else
		fsql1 = "select case_stat"
		fsql1 = fsql1 & " from case_exp where case_no='"& case_no &"'"
		rsbr1.open fsql1,conn,1,1
		if not rsbr1.eof then
			isql = "select * from exp"
			isql = isql & " where seq='"& seq &"' and seq1='"& seq1 &"'"
			rsbr2.Open isql,conn,1,1
			atp_no = trim(rsbr2("tp_no"))
			rsbr2.Close 
			if atp_no<>empty and atp_no <>"0" then
				fsql = fsql & "'LR_O',"
			else
				if rsbr1("case_stat")="N" then
					fsql = fsql & "'LR_N',"
				elseif rsbr1("case_stat")="S" then
					fsql = fsql & "'LR_N',"
				else
					fsql = fsql & "'LR_O',"
				end if
			end if
		else
			fsql = fsql & "'LR_O',"
		end if
		rsbr1.close
	end if
	fsql = fsql & "'NN','"& prgid &"',"& chknumzero(ptodosqlno)
	fsql = fsql & ")"
	'Response.Write "========== fexp:todo_exp <br>" & fsql & "<br>"
	'Response.End 
	connf.execute fsql
end function

'代理人檢核
function insert_agent_tqcA()
	dim fsql
	intofld = "insert into agent_tqc(branch,dept,in_syscode,in_prgid,seq,seq1"
	intofld = intofld & ",br_no,br_no1,agent_no,agent_no1,rs_sqlno,brs_sqlno,exch_sqlno,que_sqlno"
	intofld = intofld & ",tqc_type,tqc_item,tqc_opt,tqc_remark,in_date,in_scode,tran_date,tran_scode)"
	intofldvalue = " values('"& trim(session("se_branch")) &"','PE','"& trim(session("syscode")) &"'"
	intofldvalue = intofldvalue & ",'"& trim(prgid) & "'," & chknumzero(request("tp_no")) &",'"& trim(request("tp_no1")) & "'"
	intofldvalue = intofldvalue & ","& chknumzero(request("seq")) &",'"& trim(request("seq1")) &"'"
	intofldvalue = intofldvalue & ",'"& trim(request("expagt_no")) &"','"& trim(request("expagt_no1")) &"'"
	intofldvalue = intofldvalue & ","& chknumzero(frs_sqlno) &","& chknumzero(rs_sqlno)
	intofldvalue = intofldvalue & ","& chknumzero(request("exch_sqlno")) &","& chknumzero(que_sqlno)
	intofldvalue2 = ",getdate(),'"& trim(session("scode")) &"',getdate(),'"& trim(session("scode")) &"'"
	intofldvalue2 = intofldvalue2 & ")"
	'請款vs報價
	for k = 1 to cdbl(request("atqc_item_cnt"))
		'response.write "atqc_item:"&request("atqc_item"&k) & "<BR>"
		if trim(request("atqc_item"&k))<>empty then
			fsql = intofld & intofldvalue
			fsql = fsql & ",'A','"& trim(request("atqc_item"&k)) &"','','"& trim(request("atqc_itemremark"&k)) &"'"
			fsql = fsql & intofldvalue2
			'Response.Write fsql & "<BR>"
			connsif.Execute fsql
		end if
	next
end function
function insert_agent_tqc()
	dim fsql
	intofld = "insert into agent_tqc(branch,dept,in_syscode,in_prgid,seq,seq1"
	intofld = intofld & ",br_no,br_no1,agent_no,agent_no1,rs_sqlno,brs_sqlno,exch_sqlno,que_sqlno"
	intofld = intofld & ",tqc_type,tqc_item,tqc_opt,tqc_remark,in_date,in_scode,tran_date,tran_scode)"
	intofldvalue = " values('"& trim(session("se_branch")) &"','PE','"& trim(session("syscode")) &"'"
	intofldvalue = intofldvalue & ",'"& trim(prgid) & "'," & chknumzero(request("tp_no")) &",'"& trim(request("tp_no1")) & "'"
	intofldvalue = intofldvalue & ","& chknumzero(request("seq")) &",'"& trim(request("seq1")) &"'"
	intofldvalue = intofldvalue & ",'"& trim(request("expagt_no")) &"','"& trim(request("expagt_no1")) &"'"
	intofldvalue = intofldvalue & ","& chknumzero(frs_sqlno) &","& chknumzero(rs_sqlno)
	intofldvalue = intofldvalue & ","& chknumzero(request("exch_sqlno")) &","& chknumzero(que_sqlno)
	intofldvalue2 = ",getdate(),'"& trim(session("scode")) &"',getdate(),'"& trim(session("scode")) &"'"
	intofldvalue2 = intofldvalue2 & ")"
	'速度
	for k = 1 to cdbl(request("btqc_item_cnt"))
		'Response.Write "btqc_item: "&trim(request("btqc_item"&k)) & "<BR>"
		if trim(request("btqc_item"&k))<>empty then
			fsql = intofld & intofldvalue
			fsql = fsql & ",'B','"& trim(request("btqc_item"&k)) &"','',''"
			fsql = fsql & intofldvalue2
			'Response.Write fsql & "<BR>"
			connsif.Execute fsql
		end if
	next
	'品質
	for k = 1 to cdbl(request("ctqc_item_cnt"))
		if trim(request("ctqc_item"&k))<>empty then
			fsql = intofld & intofldvalue
			fsql = fsql & ",'C','"& trim(request("ctqc_item"&k)) &"','','"& trim(request("ctqc_itemremark"&k)) &"'"
			fsql = fsql & intofldvalue2
			'Response.Write fsql & "<BR>"
			connsif.Execute fsql
		end if
	next
end function
'代理人檢核 入log後delete
function delete_agent_tqcA()
	dim fsql
	'請款vs報價
	for k = 1 to cdbl(request("atqc_item_cnt"))
		if trim(request("atqc_sqlno"&k))<>empty then
			call insert_log_table(connsif,"U",prgid,"agent_tqc","tqc_sqlno",trim(request("atqc_sqlno"&k)))
			fsql = "delete from agent_tqc"
			'fsql = fsql & " where tqc_sqlno='"& trim(request("atqc_sqlno"&k)) &"'"
			fsql = fsql & " where branch='"& session("se_branch") &"'"
			fsql = fsql & " and br_no="& request("seq") &" and br_no1='"& request("seq1") &"'"
			fsql = fsql & " and rs_sqlno="& frs_sqlno &" and tqc_type='A'"
			'Response.Write fsql & "<BR>"
			connsif.Execute fsql
		end if
	next
end function
'代理人檢核 入log後delete
function delete_agent_tqc()
	dim fsql
	'速度
	for k = 1 to cdbl(request("btqc_item_cnt"))
		if trim(request("btqc_sqlno"&k))<>empty then
			call insert_log_table(connsif,"U",prgid,"agent_tqc","tqc_sqlno",trim(request("btqc_sqlno"&k)))
			fsql = "delete from agent_tqc where tqc_sqlno='"& trim(request("btqc_sqlno"&k)) &"'"
			'Response.Write fsql & "<BR>"
			connsif.Execute fsql
		end if
	next
	'品質
	for k = 1 to cdbl(request("ctqc_item_cnt"))
		if trim(request("ctqc_sqlno"&k))<>empty then
			call insert_log_table(connsif,"U",prgid,"agent_tqc","tqc_sqlno",trim(request("ctqc_sqlno"&k)))
			fsql = "delete from agent_tqc where tqc_sqlno='"& trim(request("ctqc_sqlno"&k)) &"'"
			'Response.Write fsql & "<BR>"
			connsif.Execute fsql
		end if
	next
end function
function update_agent_tqc_xxx()
	usqltqc = "update agent_tqc set tran_date=getdate(),tran_scode='"& trim(session("scode")) &"'"
	'速度
	for k = 1 to cdbl(request("btqc_item_cnt"))
		'Response.Write "btqc_item: "&trim(request("btqc_item"&k)) & "<BR>"
		if trim(request("btqc_item"&k))<>empty then
			usql = usqltqc
			usql = usql & ",tqc_item='"& trim(request("btqc_item"&k)) &"'"
			usql = usql & " where tqc_sqlno="& trim(request("btqc_sqlno"&k))
			'Response.Write usql & "<BR>"
			connsif.Execute usql
		end if
	next
	'請款vs報價
	for k = 1 to cdbl(request("atqc_item_cnt"))
		if trim(request("atqc_item"&k))<>empty then
			usql = usqltqc
			usql = usql & ",tqc_item='"& trim(request("atqc_item"&k)) &"'"
			usql = usql & ",tqc_remark='"& trim(request("atqc_itemremark"&k)) &"'"
			usql = usql & " where tqc_sqlno="& trim(request("atqc_sqlno"&k))
			'Response.Write usql & "<BR>"
			connsif.Execute usql
		end if
	next
	'品質
	for k = 1 to cdbl(request("ctqc_item_cnt"))
		if trim(request("ctqc_item"&k))<>empty then
			usql = usqltqc
			usql = usql & ",tqc_item='"& trim(request("ctqc_item"&k)) &"'"
			usql = usql & ",tqc_remark='"& trim(request("ctqc_itemremark"&k)) &"'"
			usql = usql & " where tqc_sqlno="& trim(request("ctqc_sqlno"&k))
			'Response.Write usql & "<BR>"
			connsif.Execute usql
		end if
	next
end function
%>
