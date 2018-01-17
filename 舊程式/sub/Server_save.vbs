<%
'-----exp, exp2, exp_mark, exp_tec, ap_exp, ap_expant, exp_attach
'--新增案件主檔
function insert_exp()
	Set rsf = Server.CreateObject("ADODB.Recordset")
	dim fsql
	fsql = "select max(exp_sqlno) + 1 from exp"
	rsf.open fsql,conn,1,1
	exp_sqlno = rsf(0)
	rsf.close
	
	fsql = "insert into exp(seq,seq1,exp_sqlno,country,case1,case_kind,ap_level,in_date," 
	fsql = fsql & "cappl_name,eappl_name,jappl_name,custprod_no,cust_prod,epctcoun_code,epctcountry,"
	fsql = fsql & "cust_area,cust_seq,att_sql,scode1,expagt_no,expagt_no1,expagent_mark,"
	fsql = fsql & "apply_date,baseapply_date,apply_no,cntry_date,pct_no,pub_date,pub_no,tapply_no,tbranch,tseq,tseq1,"
	fsql = fsql & "change_date,change_no,change_case1,appr_date,open_date,open_no,"
	fsql = fsql & "issue_date,issue_no,issue_no2,term1,term2,priority,gpay_flag,transfer_flag,"
	fsql = fsql & "ann_end_code,tp_no,tp_no1,pay_times,pay_date,step_grade,"
	fsql = fsql & "arcase_cgrs,arcase_type,arcase_class,arcase_rs_code,arcase_act_code,"
	fsql = fsql & "now_grade,now_stat,now_arcase_type,now_arcase_class,now_arcase,now_act_code,"
	fsql = fsql & "in_scode1,tr_scode,tr_date,same_apply,same_first,same_seq,same_seq1,"
	fsql = fsql & "pic_scode,pay_es,am_apply_date,am_ig,case_source,pr_scode,pr_scodee,"
	fsql = fsql & "tapply_date,rent,announce_flag,announce_date,announce_cer,"
	fsql = fsql & "seu_issue_no,eeu_issue_no,entity_flag,inspect_flag,tran_scode,tran_date"
	fsql = fsql & ",seq1z_flag,zold_seq,zold_seq1,agent_rno,prior_no"
	fsql = fsql & ")"
	fsql = fsql & " values('"& seq &"','"& seq1 &"',"& chknumzero(exp_sqlno) &","
	fsql = fsql & chkcharnull(request("country")) &","
	fsql = fsql & chkcharnull(request("case1")) &","& chkcharnull(request("case_kind")) &","
	fsql = fsql & chkcharnull(request("ap_level")) &"," 
	fsql = fsql & "'"& date() &"',"& chkcharnull2(request("cappl_name")) &","
	fsql = fsql & chkcharnull2(request("eappl_name")) &","& chkcharnull2(request("jappl_name")) &","
	fsql = fsql & chkcharnull2(request("custprod_no")) &","& chkcharnull2(request("cust_prod")) &","
	fsql = fsql & chkcharnull(request("epctcoun_code")) &","& chkcharnull(request("epctcountry")) &"," 
	fsql = fsql & chkcharnull(request("cust_area")) &","
	fsql = fsql & chknumzero(request("cust_seq")) &","& chknumzero(request("att_sql")) &","
	fsql = fsql & chkcharnull(request("mscode1")) &","& chkcharnull(request("expagt_no")) &","
	fsql = fsql & chkcharnull(request("expagt_no1")) &","& chkcharnull(expagent_mark) &","
	fsql = fsql & chkdatenull(request("apply_date")) &","& chkdatenull(request("baseapply_date")) &","& chkcharnull(request("apply_no")) &","
	fsql = fsql & chkdatenull(request("cntry_date")) &","& chkcharnull(request("pct_no")) &","
	fsql = fsql & chkdatenull(request("pub_date")) &","& chkcharnull(request("pub_no")) &","
	fsql = fsql & chkcharnull(request("tapply_no")) &","& chkcharnull(request("tbranch")) &","
	fsql = fsql & chknumzero(request("tseq")) &","& chkcharnull(request("tseq1")) &","
	fsql = fsql & chkdatenull(request("change_date")) &","& chkcharnull(request("change_no")) &","
	fsql = fsql & chkcharnull(request("change_case1")) &","& chkdatenull(request("appr_date")) &","
	fsql = fsql & chkdatenull(request("open_date")) &","& chkcharnull(request("open_no")) &","
	fsql = fsql & chkdatenull(request("issue_date")) &","& chkcharnull(request("issue_no")) &","
	fsql = fsql & chkcharnull(request("issue_no2")) &","
	fsql = fsql & chkdatenull(request("term1")) &","& chkdatenull(request("term2")) &","
	fsql = fsql & chkcharnull(priority) &","& chkcharnull(gpay_flag) &","
	fsql = fsql & chkcharnull(transfer_flag) &","& chkcharnull(request("ann_end_code")) &","
	fsql = fsql & chknumzero(request("tp_no")) &","& chkcharnull(request("tp_no1")) &","
	fsql = fsql & chkcharnull(request("pay_times")) &","& chkdatenull(request("pay_date")) &","
	fsql = fsql & chknumzero(request("step_grade")) &","
	fsql = fsql & chkcharnull(request("arcase_cgrs")) &","& chkcharnull(request("arcase_type")) &","
	fsql = fsql & chkcharnull(request("arcase_class")) &","& chkcharnull(request("arcase_rs_code")) &","
	fsql = fsql & chkcharnull(request("arcase_act_code")) &","
	fsql = fsql & chknumzero(request("now_grade")) &","& chkcharnull(request("now_stat")) &","
	fsql = fsql & chkcharnull(request("now_arcase_type")) &","& chkcharnull(request("now_arcase_class")) &","
	fsql = fsql & chkcharnull(request("now_arcase")) &","& chkcharnull(request("now_act_code")) &","
	fsql = fsql & "'"& session("scode") &"','"& session("scode") &"',"
	fsql = fsql & "getdate(),"
	fsql = fsql & chkcharnull(request("same_apply")) &","& chkcharnull(request("same_first")) &","
	fsql = fsql & chknumzero(request("same_seq")) &","
	if trim(request("same_seq1"))<>empty then
		fsql = fsql & chkcharnull(request("same_seq1")) &","
	else
		fsql = fsql & "'_',"
	end if
	fsql = fsql & chkcharnull(request("mpic_scode")) &","
	fsql = fsql & chkcharnull(request("pay_es")) &","& chkdatenull(request("am_apply_date")) &","
	fsql = fsql & chkcharnull(request("am_ig")) &","& chkcharnull(request("case_source")) &","
	fsql = fsql & chkcharnull(request("mpr_scode")) &"," & chkcharnull(request("mpr_scodee")) &","
	fsql = fsql & chkdatenull(request("tapply_date")) &","
	fsql = fsql & chkcharnull(request("rent")) &","& chkcharnull(request("announce_flag")) &","
	fsql = fsql & chkdatenull(request("announce_date")) &","& chkcharnull(request("announce_cer")) &","
	fsql = fsql & chkcharnull(request("seu_issue_no")) &","& chkcharnull(request("eeu_issue_no")) &","
	fsql = fsql & chkcharnull(request("entity_flag")) &","& chkcharnull(request("inspect_flag")) &","
	fsql = fsql & "'"& session("scode") &"',getdate()"
	fsql = fsql & ","& chkcharnull(request("seq1z_flag"))
	fsql = fsql & ","& chknumzero(request("zold_seq")) &","& chkcharnull(request("zold_seq1"))
	fsql = fsql & ","& chkcharnull(request("agent_rno")) &","& chkcharnull(request("prior_no"))
	fsql = fsql & ")"
	'Response.Write "新增案件主檔 table:exp <br>" & fsql & "<br>"
	'Response.End 
	conn.execute fsql

end function
'更新案件主檔
function update_exp()
	dim fsql
	if request("exp_sqlno")<>empty then
		call insert_log_table(conn,"U",prgid,"exp","exp_sqlno",request("exp_sqlno"))
	else
		call insert_log_table(conn,"U",prgid,"exp","seq;seq1",seq&";"&seq1)
	end if
	
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = " select * from exp"
	if request("exp_sqlno")<>empty then
		fsql = fsql & " where exp_sqlno = " & request("exp_sqlno")
	else
		fsql = fsql & " where seq = " & request("seq") & " and seq1 = '" & request("seq1") & "'"
	end if
	rsf.Open fsql,conn,1,1
	
	fsql = "update exp set tr_date=getdate(),tr_scode='"& session("scode") &"'"
	fsql = fsql & ",tran_date=getdate(),tran_scode='"& session("scode") &"'"
	fsql = fsql & ",country="& chkcharnull(request("country")) &",case1="& chkcharnull(request("case1")) 
	data_check_rec_log "char",rsf("country"),request("country"),"exp","country",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("case1"),request("case1"),"exp","case1",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",case_kind="& chkcharnull(request("case_kind")) & ",cappl_name=" & chkcharnull2(request("cappl_name")) 
	data_check_rec_log "char",rsf("case_kind"),request("case_kind"),"exp","case_kind",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("cappl_name"),request("cappl_name"),"exp","cappl_name",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",ap_level="& chkcharnull(request("ap_level")) 
	data_check_rec_log "char",rsf("ap_level"),request("ap_level"),"exp","ap_level",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",eappl_name=" & chkcharnull2(request("eappl_name")) & ",jappl_name=" & chkcharnull2(request("jappl_name")) 
	data_check_rec_log "char",rsf("eappl_name"),request("eappl_name"),"exp","eappl_name",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("jappl_name"),request("jappl_name"),"exp","jappl_name",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",custprod_no=" & chkcharnull2(request("custprod_no")) &",cust_prod=" & chkcharnull2(request("cust_prod")) 
	data_check_rec_log "char",rsf("custprod_no"),request("custprod_no"),"exp","custprod_no",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("cust_prod"),request("cust_prod"),"exp","cust_prod",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",epctcoun_code=" & chkcharnull(request("epctcoun_code"))
	data_check_rec_log "char",rsf("epctcoun_code"),request("epctcoun_code"),"exp","epctcoun_code",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",epctcountry=" & chkcharnull(request("epctcountry")) &",cust_area="& chkcharnull(request("cust_area")) 
	data_check_rec_log "char",rsf("epctcountry"),request("epctcountry"),"exp","epctcountry",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("cust_area"),request("cust_area"),"exp","cust_area",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",cust_seq=" & chknumzero(request("cust_seq")) &",att_sql="& chknumzero(request("att_sql")) 
	data_check_rec_log "int",rsf("cust_seq"),request("cust_seq"),"exp","cust_seq",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("att_sql"),request("att_sql"),"exp","att_sql",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",scode1=" & chkcharnull(request("mscode1")) &",expagt_no="& chkcharnull(request("expagt_no"))
	data_check_rec_log "char",rsf("scode1"),request("mscode1"),"exp","scode1",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("expagt_no"),request("expagt_no"),"exp","expagt_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",expagt_no1=" & chkcharnull(request("expagt_no1")) &",expagent_mark="& chkcharnull(expagent_mark) 
	data_check_rec_log "char",rsf("expagt_no1"),request("expagt_no1"),"exp","expagt_no1",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("expagent_mark"),expagent_mark,"exp","expagent_mark",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",apply_date=" & chkdatenull(request("apply_date")) &",apply_no="& chkcharnull(request("apply_no")) 
	data_check_rec_log "date",rsf("apply_date"),request("apply_date"),"exp","apply_date",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("apply_no"),request("apply_no"),"exp","apply_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",pub_date=" & chkdatenull(request("pub_date")) &",pub_no="& chkcharnull(request("pub_no")) 
	data_check_rec_log "date",rsf("pub_date"),request("pub_date"),"exp","pub_date",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("pub_no"),request("pub_no"),"exp","pub_no",request("seq"),request("seq1"),request("rs_sqlno")
	'fsql = fsql & ",tapply_no=" & chkcharnull(request("tapply_no")) &",tbranch="& chkcharnull(request("tbranch"))
	'data_check_rec_log "char",rsf("tapply_no"),request("tapply_no"),"exp","tapply_no",request("seq"),request("seq1"),request("rs_sqlno")
	'data_check_rec_log "char",rsf("tbranch"),request("tbranch"),"exp","tbranch",request("seq"),request("seq1"),request("rs_sqlno")
	'fsql = fsql & ",tseq=" & chknumzero(request("tseq")) &",tseq1="& chkcharnull(request("tseq1"))
	'data_check_rec_log "int",rsf("tseq"),request("tseq"),"exp","tseq",request("seq"),request("seq1"),request("rs_sqlno")
	'data_check_rec_log "char",rsf("tseq1"),request("tseq1"),"exp","tseq1",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",change_date=" & chkdatenull(request("change_date")) &",change_no="& chkcharnull(request("change_no")) 
	data_check_rec_log "date",rsf("change_date"),request("change_date"),"exp","change_date",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("change_no"),request("change_no"),"exp","change_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",change_case1=" & chkcharnull(request("change_case1"))
	data_check_rec_log "char",rsf("change_case1"),request("change_case1"),"exp","change_case1",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",open_date=" & chkdatenull(request("open_date")) &",open_no="& chkcharnull(request("open_no"))
	data_check_rec_log "date",rsf("open_date"),request("open_date"),"exp","open_date",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("open_no"),request("open_no"),"exp","open_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",issue_date=" & chkdatenull(request("issue_date")) &",issue_no="& chkcharnull(request("issue_no")) 
	data_check_rec_log "date",rsf("issue_date"),request("issue_date"),"exp","issue_date",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("issue_no"),request("issue_no"),"exp","issue_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",issue_no2=" & chkcharnull(request("issue_no2")) 
	data_check_rec_log "char",rsf("issue_no2"),request("issue_no2"),"exp","issue_no2",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",term1=" & chkdatenull(request("term1")) &",term2="& chkdatenull(request("term2")) 
	data_check_rec_log "date",rsf("term1"),request("term1"),"exp","term1",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "date",rsf("term2"),request("term2"),"exp","term2",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",appr_date=" & chkdatenull(request("appr_date")) 
	data_check_rec_log "date",rsf("appr_date"),request("appr_date"),"exp","appr_date",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",priority=" & chkcharnull(priority) &",gpay_flag="& chkcharnull(gpay_flag)
	data_check_rec_log "char",rsf("priority"),priority,"exp","priority",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("gpay_flag"),gpay_flag,"exp","gpay_flag",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",transfer_flag=" & chkcharnull(transfer_flag) &",ann_end_code="& chkcharnull(request("ann_end_code")) 
	data_check_rec_log "date",rsf("transfer_flag"),transfer_flag,"exp","term1",request("transfer_flag"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "date",rsf("ann_end_code"),request("ann_end_code"),"exp","term1",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",tp_no=" & chknumzero(request("tp_no")) & ",tp_no1=" & chkcharnull(request("tp_no1"))  
	data_check_rec_log "int",rsf("tp_no"),request("tp_no"),"exp","tp_no",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("tp_no1"),request("tp_no1"),"exp","tp_no1",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",pay_times=" & chkcharnull(request("pay_times")) &",pay_date="& chkdatenull(request("pay_date")) 
	data_check_rec_log "char",rsf("pay_times"),request("pay_times"),"exp","pay_times",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "date",rsf("pay_date"),request("pay_date"),"exp","pay_date",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	if request("nstep_grade") = "1" then
		if prgid<>"exp33" then
			fsql = fsql & ",step_grade=" & chknumzero(request("nstep_grade"))
			data_check_rec_log "int",rsf("step_grade"),request("nstep_grade"),"exp","step_grade",request("seq"),request("seq1"),request("rs_sqlno")
		end if
		fsql = fsql & ",arcase_cgrs=" & chkcharnull(request("arcase_cg")) &",arcase_type="& chkcharnull(request("arcase_type")) 
		data_check_rec_log "char",rsf("arcase_cgrs"),request("arcase_cgrs"),"exp","arcase_cgrs",request("seq"),request("seq1"),request("rs_sqlno")
		data_check_rec_log "char",rsf("arcase_type"),request("arcase_type"),"exp","arcase_type",request("seq"),request("seq1"),request("rs_sqlno")
		fsql = fsql & ",arcase_class=" & chkcharnull(request("arcase_class")) &",arcase_rs_code="& chkcharnull(request("arcase_rs_code")) 
		data_check_rec_log "char",rsf("arcase_class"),request("arcase_class"),"exp","arcase_class",request("seq"),request("seq1"),request("rs_sqlno")
		data_check_rec_log "char",rsf("arcase_rs_code"),request("arcase_rs_code"),"exp","arcase_rs_code",request("seq"),request("seq1"),request("rs_sqlno")
		fsql = fsql & ",arcase_act_code=" & chkcharnull(request("arcase_act_code"))
		data_check_rec_log "char",rsf("arcase_act_code"),request("arcase_act_code"),"exp","arcase_act_code",request("seq"),request("seq1"),request("rs_sqlno")
	else
		if trim(request("ncase_stat"))<>empty then
			fsql = fsql & ",now_grade=" & chknumzero(request("now_grade")) &",now_stat="& chkcharnull(request("now_stat")) 
			data_check_rec_log "int",rsf("now_grade"),request("now_grade"),"exp","now_grade",request("seq"),request("seq1"),request("rs_sqlno")
			data_check_rec_log "char",rsf("now_stat"),request("now_stat"),"exp","now_stat",request("seq"),request("seq1"),request("rs_sqlno")
			fsql = fsql & ",now_arcase_type=" & chkcharnull(request("now_arcase_type")) &",now_arcase_class="& chkcharnull(request("now_arcase_class")) 
			data_check_rec_log "char",rsf("now_arcase_type"),request("now_arcase_type"),"exp","now_arcase_type",request("seq"),request("seq1"),request("rs_sqlno")
			data_check_rec_log "char",rsf("now_arcase_class"),request("now_arcase_class"),"exp","now_arcase_class",request("seq"),request("seq1"),request("rs_sqlno")
			fsql = fsql & ",now_arcase=" & chkcharnull(request("now_arcase")) &",now_act_code="& chkcharnull(request("now_act_code")) 
			data_check_rec_log "char",rsf("now_arcase"),request("now_arcase"),"exp","now_arcase",request("seq"),request("seq1"),request("rs_sqlno")
			data_check_rec_log "char",rsf("now_act_code"),request("now_act_code"),"exp","now_act_code",request("seq"),request("seq1"),request("rs_sqlno")
		end if
	end if
	'同一內容同日申請發明與新型
	fsql = fsql & ",same_apply="& chkcharnull(request("same_apply")) 
	if trim(request("same_apply"))="N" then
		fsql = fsql & ",same_first='',same_seq=0,same_seq1=''"
	else
		fsql = fsql & ",same_first="& chkcharnull(request("same_first")) 
		if trim(request("same_first"))="Y" then
			'此案件為前案
			fsql = fsql & ",same_seq=0,same_seq1=''"
		else
			'此案件為後案
			fsql = fsql & ",same_seq=" & chknumzero(request("same_seq")) 
			if trim(request("same_seq1"))<>empty then
				fsql = fsql & ",same_seq1=" & chkcharnull(request("same_seq1"))  
			else
				fsql = fsql & ",same_seq1='_'"
			end if
		end if
	end if
	data_check_rec_log "char",rsf("same_apply"),request("same_apply"),"exp","same_apply",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("same_first"),request("same_first"),"exp","same_first",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "int",rsf("same_seq"),request("same_seq"),"exp","same_seq",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("same_seq1"),request("same_seq1"),"exp","same_seq1",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",pic_scode="& chkcharnull(request("mpic_scode")) 
	data_check_rec_log "char",rsf("pic_scode"),request("mpic_scode"),"exp","pic_scode",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",pay_es=" & chkcharnull(request("pay_es")) &",am_apply_date="& chkdatenull(request("am_apply_date")) 
	data_check_rec_log "char",rsf("pay_es"),request("pay_es"),"exp","pay_es",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "date",rsf("am_apply_date"),request("am_apply_date"),"exp","am_apply_date",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",am_ig=" & chkcharnull(request("am_ig")) &",case_source="& chkcharnull(request("case_source")) 
	data_check_rec_log "char",rsf("am_ig"),request("am_ig"),"exp","am_ig",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("case_source"),request("case_source"),"exp","case_source",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",pr_scode=" & chkcharnull(request("mpr_scode")) 
	data_check_rec_log "char",rsf("pr_scode"),request("mpr_scode"),"exp","pr_scode",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",pr_scodee=" & chkcharnull(request("mpr_scodee")) 
	data_check_rec_log "char",rsf("pr_scodee"),request("mpr_scodee"),"exp","pr_scodee",request("seq"),request("seq1"),request("rs_sqlno")
	'fsql = fsql & ",tapply_date="& chkdatenull(request("tapply_date")) 
	'data_check_rec_log "date",rsf("tapply_date"),request("tapply_date"),"exp","tapply_date",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",rent=" & chkcharnull(request("rent")) &",announce_flag="& chkcharnull(request("announce_flag")) 
	data_check_rec_log "char",rsf("rent"),request("rent"),"exp","rent",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("announce_flag"),request("announce_flag"),"exp","announce_flag",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",announce_date=" & chkdatenull(request("announce_date")) &",announce_cer="& chkcharnull(request("announce_cer")) 
	data_check_rec_log "date",rsf("announce_date"),request("announce_date"),"exp","announce_date",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("announce_cer"),request("announce_cer"),"exp","announce_cer",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",seu_issue_no=" & chkcharnull(request("seu_issue_no")) 
	data_check_rec_log "char",rsf("seu_issue_no"),request("seu_issue_no"),"exp","seu_issue_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",eeu_issue_no=" & chkcharnull(request("eeu_issue_no")) 
	data_check_rec_log "char",rsf("eeu_issue_no"),request("eeu_issue_no"),"exp","eeu_issue_no",request("seq"),request("seq1"),request("rs_sqlno")
	'fsql = fsql & ",entity_flag=" & chkcharnull(request("entity_flag")) &""
	'data_check_rec_log "char",rsf("entity_flag"),request("entity_flag"),"exp","entity_flag",request("seq"),request("seq1"),request("rs_sqlno")
	'fsql = fsql & ",inspect_flag="& chkcharnull(request("inspect_flag"))
	'data_check_rec_log "char",rsf("inspect_flag"),request("inspect_flag"),"exp","inspect_flag",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",end_date="& chkdatenull(request("end_date")) &",end_code="& chkcharnull(request("end_code"))
	fsql = fsql & ",endremark="& chkdatenull(request("endremark"))
	data_check_rec_log "date",rsf("end_date"),request("end_date"),"exp","end_date",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("end_code"),request("end_code"),"exp","end_code",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("endremark"),request("endremark"),"exp","endremark",request("seq"),request("seq1"),request("rs_sqlno")
	'fsql = fsql & ",t_end_date="& chkdatenull(request("t_end_date")) 
	'data_check_rec_log "date",rsf("t_end_date"),request("t_end_date"),"exp","t_end_date",request("ann_end_code"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",seq1z_flag="& chkcharnull(request("seq1z_flag")) & ",zold_flag="& chkcharnull(request("zold_flag"))
	fsql = fsql & ",zold_seq="& chknumzero(request("zold_seq")) &",zold_seq1="& chkcharnull(request("zold_seq1"))
	data_check_rec_log "char",rsf("seq1z_flag"),request("seq1z_flag"),"exp","seq1z_flag",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("zold_flag"),request("zold_flag"),"exp","zold_flag",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "int",rsf("zold_seq"),request("zold_seq"),"exp","zold_seq",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("zold_seq1"),request("zold_seq1"),"exp","zold_seq1",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",baseapply_date="& chkdatenull(request("baseapply_date"))
	data_check_rec_log "date",rsf("baseapply_date"),request("baseapply_date"),"exp","baseapply_date",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",cntry_date="& chkdatenull(request("cntry_date")) & ",pct_no="& chkcharnull(request("pct_no"))
	data_check_rec_log "date",rsf("cntry_date"),request("cntry_date"),"exp","cntry_date",request("seq"),request("seq1"),request("rs_sqlno")
	data_check_rec_log "char",rsf("pct_no"),request("pct_no"),"exp","pct_no",request("seq"),request("seq1"),request("rs_sqlno")
	fsql = fsql & ",agent_rno="& chkcharnull(request("agent_rno"))
	data_check_rec_log "char",rsf("agent_rno"),request("agent_rno"),"exp","agent_rno",request("seq"),request("seq1"),request("rs_sqlno")	
	fsql = fsql & ",prior_no="& chkcharnull(request("prior_no"))
	data_check_rec_log "char",rsf("prior_no"),request("prior_no"),"exp","prior_no",request("seq"),request("seq1"),request("rs_sqlno")	
	if request("exp_sqlno")<>empty then
		fsql = fsql & " where exp_sqlno = " & request("exp_sqlno")
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "更新案件檔 table:exp <br>" & fsql & "<br>"
	'response.end
	conn.execute fsql
	
	'分案主檔案件名稱
	if request("seq")<>empty then
		fsql = "update br_exp set cappl_name="& chkcharnull2(request("cappl_name")) 
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
		'Response.Write "更新案件檔 table:exp <br>" & fsql & "<br>"
		'response.end
		conn.execute fsql
	end if
end function
function data_check_rec_log(kind,ovalue,nvalue,ptable,fldname,pseq,pseq1,prs_sqlno)
	select case kind
		case "char"
			if trim(ovalue) <> trim(nvalue) then 
				insert_exp_rec_log conn,pseq,pseq1,ptable,fldname,ovalue,nvalue,prgid,prs_sqlno
			end if
		case "date","int"
			if trim(ovalue) <> trim(nvalue) or (isnull(trim(ovalue)) and trim(nvalue) <> "") or (trim(ovalue) <> empty and trim(nvalue) = "") then
				insert_exp_rec_log conn,pseq,pseq1,ptable,fldname,ovalue,nvalue,prgid,prs_sqlno
			end if
	end select
end function 
'刪除案件主檔
function delete_exp()
	dim fsql
	if exp_sqlno<>empty then
		call insert_log_table(conn,"D",prgid,"exp","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"D",prgid,"exp","seq;seq1",seq&";"&seq1)
	end if
	fsql = "delete from exp"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "刪除案件主檔 table:exp <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'備註
function update_exp_mark()
	dim i
	dim fsql
	dim rsf
	SET rsf = server.CreateObject("ADODB.RECORDSET")
	fsql = "select * from exp_mark"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'response.write fsql & "<BR>"
	rsf.open fsql,conn,1,1
	if not rsf.eof then
		if exp_sqlno<>empty then
			call insert_log_table(conn,"U",prgid,"exp_mark","exp_sqlno",exp_sqlno)
		else
			call insert_log_table(conn,"U",prgid,"exp_mark","seq;seq1",seq&";"&seq1)
		end if
		fsql = "update exp_mark set slang="& chkcharnull(request("slang")) &",spages="& chknumzero(request("spages"))
		fsql = fsql & ",words="& chknumzero(request("words")) &",claim="& chknumzero(request("claim"))
		fsql = fsql & ",draw_cnt="& chknumzero(request("draw_cnt")) &",draw_paper="& chkcharnull(request("draw_paper"))
		fsql = fsql & ",ipc="& chkcharnull(request("ipc")) &",ind_item="& chknumzero(request("ind_item"))
		fsql = fsql & ",copy_flag="& chkcharnull(request("copy_flag"))
		if request("copy_flag")="Y" then
		    fsql = fsql & ",copy_seq="& chknumzero(request("copy_seq"))
		    fsql = fsql & ",copy_seq1="& chkcharnull(request("copy_seq1"))
		else
		    fsql = fsql & ",copy_seq=0,copy_seq1=''"
		end if
		fsql = fsql & ",family_master="& chkcharnull(request("family_master"))
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
		fsql = fsql & ",hk_flag="& chkcharnull(request("hk_flag"))
		fsql = fsql & ",hk_seq="& chknumzero(request("hk_seq")) & ",hk_seq1="& chkcharnull(request("hk_seq1"))
		fsql = fsql & ",eu_seq="& chknumzero(request("eu_seq")) & ",eu_seq1="& chkcharnull(request("eu_seq1"))
		fsql = fsql & ",fremark="& chkcharnull2(request("fremark"))
		fsql = fsql & ",bremark="& chkcharnull2(request("bremark"))
		fsql = fsql & ",tran_date=getdate(),tran_scode="& chkcharnull(session("scode"))
		if exp_sqlno<>empty then
			fsql = fsql & " where exp_sqlno = " & exp_sqlno
		else
			fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
		end if
	else
		'舊案會有沒有exp_mark的狀況，要insert一筆
		fsql = "insert into exp_mark(exp_sqlno,seq,seq1,slang,spages,words,claim,draw_cnt,draw_paper,"
		fsql = fsql & "ipc,ind_item,copy_flag,copy_seq,copy_seq1,family_master,family_dept,family_seq,family_seq1,"
		fsql = fsql & "family_flag,family_country,family_apply_no,hk_flag,hk_seq,hk_seq1,eu_seq,eu_seq1,"
		fsql = fsql & "fremark,bremark,tran_date,tran_scode)"
		fsql = fsql & " values("
		if request("exp_sqlno")<>empty then
			fsql = fsql & chknumzero(exp_sqlno) &",'"& seq &"','"& seq1 &"'"
		else
			if request("newold")="N" or request("newold")="S" then
				if request("newold")="N" then
					fsql = fsql & chknumzero(exp_sqlno) &",null,''"
				else
					fsql = fsql & chknumzero(exp_sqlno) &","& seq &",'"& seq1 &"'"
				end if
			else
				fsql = fsql & chknumzero(exp_sqlno) &","& seq &",'"& seq1 &"'"
			end if
		end if
		fsql = fsql & ","& chkcharnull(request("slang")) &","& chknumzero(request("spages"))
		fsql = fsql & ","& chknumzero(request("words")) &","& chknumzero(request("claim"))
		fsql = fsql & ","& chknumzero(request("draw_cnt")) &","& chkcharnull(request("draw_paper"))
		fsql = fsql & ","& chkcharnull(request("ipc")) &","& chknumzero(request("ind_item"))
		fsql = fsql & ","& chkcharnull(request("copy_flag")) &","& chknumzero(request("copy_seq"))
		fsql = fsql & ","& chkcharnull(request("copy_seq1"))
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
		fsql = fsql & ","& chkcharnull(request("hk_flag")) &","& chknumzero(request("hk_seq"))
		fsql = fsql & ","& chkcharnull(request("hk_seq1"))
		fsql = fsql & ","& chknumzero(request("eu_seq")) & ","& chkcharnull(request("eu_seq1"))
		fsql = fsql & ","& chkcharnull2(request("fremark")) &","& chkcharnull2(request("bremark"))
		fsql = fsql & ",getdate(),"& chkcharnull(session("scode"))
		fsql = fsql & ")"
	end if
	'if session("scode")="admin" then
	'Response.Write "更新案件備註檔 table:exp_mark <br>" & fsql & "<br>"
	'response.end 
	'end if
	conn.execute fsql
end function
'刪除案件備註
function delete_exp_mark()
	dim fsql
	if exp_sqlno<>empty then
		call insert_log_table(conn,"D",prgid,"exp_mark","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"D",prgid,"exp_mark","seq;seq1",seq&";"&seq1)
	end if
	fsql = "delete from exp_mark"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "刪除案件備註檔 table:exp_mark <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'優先權資料
function insert_exp_prior(pmustDel,pnum)
	if exp_sqlno<>empty then
		call insert_log_table(conn,"U",prgid,"exp2","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"U",prgid,"exp2","seq;seq1",seq&";"&seq1)
	end if
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from exp2"
		if exp_sqlno<>empty then
			fsql = fsql & " where exp_sqlno = " & exp_sqlno
		else
			fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
		end if
		'Response.Write "更新優先權資料 table:exp2 <br>" & fsql & "<br>"
		conn.execute fsql
	end if
	for i=1 to pnum
		if request("hpriordel_flag"&i)="Y" then
			fsql = "update exp2_log set ud_flag='D'"
			if exp_sqlno<>empty then
				fsql = fsql & " where exp_sqlno = " & exp_sqlno
			else
				fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
			end if
			fsql = fsql & " and seqno="& request("prior_sqlno"&i)
			'Response.Write "更新優先權資料 table:exp2 <br>" & fsql & "<br>"
			conn.execute fsql
		else
			fsql = "insert into exp2(country,exp_sqlno,seq,seq1,apply_date,apply_no,apply_yn"
			fsql = fsql & ",prior_yn,apply_seq,apply_seq1,prior_no,tran_date,tran_scode)"
			fsql = fsql & " values("& chkcharnull(request("prior_country"&i))
			if exp_sqlno<>empty then
				fsql = fsql & ","& chknumzero(exp_sqlno)
				if seq<>empty then
					fsql = fsql & ","& seq &",'"& seq1 &"'"
				else
					fsql = fsql & ",null,''"
				end if
			else
				if trim(request("newold"))="N" or trim(request("newold"))="S" then
					if request("newold")="N" then
						fsql = fsql & ","& chknumzero(exp_sqlno) &",null,''"
					else
						fsql = fsql & ","& chknumzero(exp_sqlno) &","& seq &",'"& seq1 &"'"
					end if
				else
					fsql = fsql & ","& chknumzero(exp_sqlno) &","& seq &",'"& seq1 &"'"
				end if
			end if
			fsql = fsql & ","& chkdatenull(request("apply_date"&i)) &","& chkcharnull(request("apply_no"&i)) 
			fsql = fsql & ","& chkcharnull(request("apply_yn"&i)) & ","& chkcharnull(request("prior_yn"&i))
			fsql = fsql & ","& chkcharnull(request("apply_seq_"&i)) & ","& chkcharnull(request("apply_seq1_"&i))  
			fsql = fsql & ","& chkcharnull(request("prior_no"&i))
			fsql = fsql & ",getdate(),'"& session("scode") &"')"
			'Response.Write "更新優先權資料 table:exp2 <br>" & fsql & "<br>"
			'response.end
			conn.execute fsql
		end if
	next
	'response.end
end function
'刪除優先權資料
function delete_exp2()
	dim fsql
	if exp_sqlno<>empty then
		call insert_log_table(conn,"D",prgid,"exp2","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"D",prgid,"exp2","seq;seq1",seq&";"&seq1)
	end if
	fsql = "delete from exp2"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "刪除優先權資料 table:exp2 <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'技術相關案件記錄檔
function insert_exp_tec(pmustDel,pnum,pflag)
	if exp_sqlno<>empty then
		call insert_log_table(conn,"U",prgid,"exp_tec","exp_sqlno;tec_flag",exp_sqlno&";"&pflag)
	else
		call insert_log_table(conn,"U",prgid,"exp_tec","seq;seq1;tec_flag",seq&";"&seq1&";"&pflag)
	end if
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from exp_tec"
		if seq<>empty and seq<>"0" then
			fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
		else
			fsql = fsql & " where exp_sqlno = " & exp_sqlno
		end if
		fsql = fsql & " and tec_flag='"& pflag &"'"
		'response.write fsql & "<BR>"
		conn.execute fsql
	end if
	for i=1 to pnum
		if pflag="brp" then
			tec_sqlno = request("brptec_sqlno"&i)
			tec_code = request("brp_code"&i)
			tec_seq = request("tec_brp_seq"&i)
			tec_seq1 = request("tec_brp_seq1_"&i)
			tec_country = "T"
			tec_apply_no = request("tec_bapply_no"&i)
			del_flag = request("hbrpdel_flag"&i)
		else
			tec_sqlno = request("exptec_sqlno"&i)
			tec_code = request("exp_code"&i)
			tec_seq = request("tec_exp_seq"&i)
			tec_seq1 = request("tec_exp_seq1_"&i)
			tec_country = request("tec_ecountry"&i)
			tec_apply_no = request("tec_eapply_no"&i)
			del_flag = request("hexpdel_flag"&i)
		end if
		if del_flag="Y" then
			fsql = "update exp_tec_log set ud_flag='D'"
			if exp_sqlno<>empty then
				fsql = fsql & " where exp_sqlno = " & exp_sqlno
			else
				fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
			end if
			fsql = fsql & " and exptec_sqlno="& tec_sqlno
			'Response.Write "更新相關案件檔 table:exp_tec <br>" & fsql & "<br>"
			conn.execute fsql
		else
			fsql = "insert into exp_tec(seq,seq1,exp_sqlno,tec_flag,tec_code,tec_seq,tec_seq1"
			fsql = fsql & ",tec_country,tec_apply_no,tran_date,tran_scode)"
			fsql = fsql & " values("
			if exp_sqlno<>empty then
				if seq<>empty then
					fsql = fsql & seq &",'"& seq1 &"',"& chknumzero(exp_sqlno)
				else
					fsql = fsql & "null,'',"& chknumzero(exp_sqlno)
				end if
			else
				if request("newold")="N" or request("newold")="S" then
					if request("newold")="N" then
						fsql = fsql & "null,'',"& chknumzero(exp_sqlno) 
					else
						fsql = fsql & seq &",'"& seq1 &"',"& chknumzero(exp_sqlno)
					end if
				else
					fsql = fsql & seq &",'"& seq1 &"',"& chknumzero(exp_sqlno)
				end if
			end if
			fsql = fsql & ","& chkcharnull(pflag) &","& chkcharnull(tec_code) 
			fsql = fsql & ","& chknumzero(tec_seq) &","& chkcharnull(tec_seq1) 
			fsql = fsql & ","& chkcharnull(tec_country) &","& chkcharnull(tec_apply_no) 
			fsql = fsql & ","& "getdate(),'"& session("scode") &"')"
			'Response.Write "更新相關案件檔 table:exp_tec <br>" & fsql & "<br>"
			'response.end
			conn.execute fsql
		end if
	next
	'response.end
end function
'刪除相關案件檔
function delete_exp_tec()
	dim fsql
	if exp_sqlno<>empty then
		call insert_log_table(conn,"D",prgid,"exp_tec","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"D",prgid,"exp_tec","seq;seq1",seq&";"&seq1)
	end if
	fsql = "delete from exp_tec"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "刪除相關案件檔 table:exp_tec <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'更新案件申請人檔
function insert_ap_exp(pmustDel,pnum)
	call insert_log_table(conn,"U",prgid,"ap_exp","seq;seq1",seq&";"&seq1)
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from ap_exp where seq='"& seq &"' and seq1='"& seq1 &"'"
		conn.execute fsql
	end if
	for i=1 to pnum
		if request("hapdel_flag"&i)="Y" then
			fsql = "update ap_exp_log set ud_flag='D'"
			fsql = fsql & " where seq="& seq &" and seq1='"& seq1 &"'"
			fsql = fsql & " and sqlno="& request("hap_sqlno"&i)
			conn.execute fsql
		else
			fsql = "insert into ap_exp(seq,seq1,exp_sqlno,apsqlno,apcust_no,ap_cname1,ap_cname2"
			fsql = fsql & ",ap_brith,ap_title,ap_crep,ap_erep,ap_sql,tran_date,tran_scode)"
			fsql = fsql & " values("& seq &",'"& seq1 &"',"& chknumzero(exp_sqlno)
			fsql = fsql & ","& chknumzero(request("apsqlno"&i)) & ","& chkcharnull(request("apcust_no"&i))
			fsql = fsql & ","& chkcharnull2(request("ap_cname1_"&i)) &","& chkcharnull2(request("ap_cname2_"&i))
			fsql = fsql & ","& chkdatenull(request("ap_brith"&i)) &","& chkcharnull(request("ap_title"&i))
			fsql = fsql & ","& chkcharnull2(request("ap_crep"&i)) &","& chkcharnull2(request("ap_erep"&i))
			fsql = fsql & ","& chknumzero(request("ap_sql"&i)) &",getdate(),'"& session("scode") &"')"
			'Response.Write "更新案件申請人檔 table:ap_exp <br>" & fsql & "<br>"
			conn.execute fsql
		end if
	next
end function
'由交辦申請人檔寫入案件申請人檔
function insert_ap_exp_case(pseq,pseq1,pexp_sqlno)
	dim fsql
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = " select a.*,b.apcust_no from exp_apcust a,apcust b"
	fsql = fsql & " where a.apsqlno=b.apsqlno and in_no = '" & in_no &"'"
	'response.write fsql & "<BR>"
	'response.end 
	rsf.Open fsql,conn,1,1
	while not rsf.eof 
		fsql = "insert into ap_exp(seq,seq1,exp_sqlno,apsqlno,apcust_no,ap_cname1,ap_cname2"
		fsql = fsql & ",ap_brith,ap_title,ap_crep,ap_erep,ap_sql,tran_date,tran_scode)"
		fsql = fsql & " values("& pseq &",'"& pseq1 &"',"& pexp_sqlno
		fsql = fsql & ","& chknumzero(rsf("apsqlno")) & ","& chkcharnull(rsf("apcust_no"))
		fsql = fsql & ","& chkcharnull2(rsf("ap_cname1")) &","& chkcharnull2(rsf("ap_cname2"))
		fsql = fsql & ","& chkdatenull(rsf("brith")) &","& chkcharnull(rsf("title"))
		fsql = fsql & ","& chkcharnull2(rsf("ap_crep")) &","& chkcharnull2(rsf("ap_erep"))
		fsql = fsql & ","& chknumzero(rsf("ap_sql")) &",getdate(),'"& session("scode") &"')"
		'Response.Write "更新案件申請人檔 table:ap_exp <br>" & fsql & "<br>"
		conn.execute fsql
		rsf.movenext
	wend
	rsf.close
	'response.end
end function
'刪除案件申請人檔
function delete_ap_exp()
	dim fsql
	if exp_sqlno<>empty then
		call insert_log_table(conn,"D",prgid,"ap_exp","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"D",prgid,"ap_exp","seq;seq1",seq&";"&seq1)
	end if
	fsql = "delete from ap_exp"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "刪除案件申請人檔 table:ap_exp <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'更新案件發明/創作人檔
function insert_ap_expant(pmustDel,pnum)
	call insert_log_table(conn,"U",prgid,"ap_expant","seq;seq1",seq&";"&seq1)
	dim i
	dim fsql
	if pmustDel="D" then
		fsql = "delete from ap_expant where seq='"& seq &"' and seq1='"& seq1 &"'"
		conn.execute fsql
	end if
	for i=1 to pnum
		if request("hantdel_flag"&i)="Y" then
			fsql = "update ap_expant_log set ud_flag='D'"
			fsql = fsql & " where seq="& seq &" and seq1='"& seq1 &"'"
			fsql = fsql & " and sqlno="& request("antsqlno"&i)
			conn.execute fsql
		else
			fsql = "insert into ap_expant(seq,seq1,exp_sqlno,antsqlno,ant_no,ant_cname1,ant_cname2"
			fsql = fsql & ",antcomp,tran_date,tran_scode,ant_id)"
			fsql = fsql & " values("& seq &",'"& seq1 &"',"& chknumzero(exp_sqlno)
			fsql = fsql & ","& chkcharnull(request("antsqlno"&i))
			fsql = fsql & ","& chkcharnull(request("ant_no"&i)) 
			fsql = fsql & ","& chkcharnull2(request("ant_cname1_"&i)) &","& chkcharnull2(request("ant_cname2_"&i))
			fsql = fsql & ","& chkcharnull2(request("antcomp"&i)) 
			fsql = fsql & ","& "getdate(),'"& session("scode") &"',"& chkcharnull(request("ant_id"&i)) &")"
			'Response.Write "更新案件發明/創作人檔 table:ap_expant <br>" & fsql & "<br>"
			conn.execute fsql
		end if
	next
end function
'由交辦申請人檔寫入案件發明/創作人檔
function insert_ap_expant_case(pseq,pseq1,pexp_sqlno)
	dim fsql
	Set rsf = Server.CreateObject("ADODB.Recordset")
	fsql = " select a.*,b.ant_no from exp_ant a,inventor b"
	fsql = fsql & " where a.antsqlno=b.antsqlno and in_no = '" & in_no &"'"
	'response.write fsql & "<BR>"
	'response.end 
	rsf.Open fsql,conn,1,1
	while not rsf.eof 
		fsql = "insert into ap_expant(seq,seq1,exp_sqlno,antsqlno,ant_no,ant_cname1,ant_cname2"
		fsql = fsql & ",antcomp,tran_date,tran_scode,ant_id)"
		fsql = fsql & " values("& pseq &",'"& pseq1 &"',"& chknumzero(pexp_sqlno)
		fsql = fsql & ","& chkcharnull(rsf("antsqlno"))
		fsql = fsql & ","& chkcharnull(rsf("ant_no")) 
		fsql = fsql & ","& chkcharnull2(rsf("ant_cname1")) &","& chkcharnull2(rsf("ant_cname2"))
		fsql = fsql & ","& chkcharnull2(rsf("antcomp")) 
		fsql = fsql & ","& "getdate(),'"& session("scode") &"',"& chkcharnull(rsf("ant_id")) &")"
		'Response.Write "更新案件發明/創作人檔 table:ap_expant <br>" & fsql & "<br>"
		conn.execute fsql
		rsf.movenext
	wend
	rsf.close
end function
'刪除案件發明/創作人檔
function delete_ap_expant()
	dim fsql
	if exp_sqlno<>empty then
		call insert_log_table(conn,"D",prgid,"ap_expant","exp_sqlno",exp_sqlno)
	else
		call insert_log_table(conn,"D",prgid,"ap_expant","seq;seq1",seq&";"&seq1)
	end if
	fsql = "delete from ap_expant"
	if exp_sqlno<>empty then
		fsql = fsql & " where exp_sqlno = " & exp_sqlno
	else
		fsql = fsql & " where seq = " & seq & " and seq1 = '" & seq1 & "'"
	end if
	'Response.Write "刪除案件發明/創作人檔 table:ap_expant <br>" & fsql & "<br>"
	conn.execute fsql
end function 
'處理上傳圖檔的部份
Function upin_exp_attach(pseq,pseq1,pstep_grade,psource,patt_sqlno)
	dim i
	'---目前資料庫中有的最大值的欄位名稱
	maxAttach_no = request("uploadfield")&"_maxAttach_no"
	'response.write  "目前資料庫中有的最大值--maxAttach_no=" & request(maxAttach_no) &"<br>"
	'---exp_attach.attach_no的欄位名稱
	filenum=request("uploadfield")&"filenum"
	'response.write  "exp_attach.attach_no--filenum=" & request(filenum) &"<br>"
	'---畫面NO顯示編號的欄位名稱
	sqlnum=request("uploadfield")&"sqlnum"
	'response.write  "畫面NO顯示編號--sqlnum=" & request(sqlnum) &"<br>"
	'---目前table的筆數的欄位名稱
	attach_cnt=request("uploadfield")&"_attach_cnt"
	'response.write "目前table的筆數--attach_cnt="& request(attach_cnt) &"<BR>"
	'---欄位名稱
	uploadfield = trim(request("uploadfield"))
	'response.write "欄位名稱--uploadfield="& uploadfield & "<BR>"
	
	'--*** Log 需另外入，因step_grade可能不一樣
	'--*** call insert_log_table(conn,"U",prgid,"exp_attach","seq;seq1;step_grade",seq&";"&seq1&";0")
	
	Select Case request(maxattach_no)
	Case "0"		'表示資料庫無資料則用新增方式
		attachno=1
		for i=1 to cint(request(filenum))
			IF trim(request(uploadfield&i))<>empty then
				if trim(request(uploadfield & "_open_flag" & i))="on" _
				or trim(request(uploadfield & "_open_flag" & i))="Y" then
					open_flag = "Y"
				else
					open_flag = "N"
				end if
				fsql = "insert into exp_attach (exp_sqlno,Seq,seq1,step_grade,case_no,att_sqlno" &_
					   ",Source,in_date,in_scode,Attach_no,attach_path,doc_type,attach_desc" &_
					   ",Attach_name,Attach_size,source_name,attach_flag,Mark,open_flag,tran_date,tran_scode) values (" &_
					   chknumzero(request("exp_sqlno")) &",'"& trim(pseq) &"','"& trim(pseq1) &"'," &_
					   chknumzero(pstep_grade) &","& chkcharnull(request("case_no")) &","
				fsql = fsql & "'"& patt_sqlno &"','"& psource &"',"
				fsql = fsql & "getdate(),'"& session("scode")&"','"& attachno &"','"& request(uploadfield&i) &"',"
				fsql = fsql & "'"& trim(request(uploadfield & "_doc_type" & i)) &"',"
				fsql = fsql & "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"'," &_
					   "'"& trim(request(uploadfield & "_size" & i)) &"','"& request(uploadfield & "_source_name" & i) &"'," &_
					   "'A','','"& open_flag &"',getdate(),'"& session("scode") &"')"
				if session("scode")="admin" then response.write "資料庫無資料新增exp_attach"& i &"<br>" &fsql&"<br><br>"
				attachno = attachno + 1
				Conn.execute fsql
			End IF	
		next
	Case else	'當資料庫已經有值
		Set rsf = Server.CreateObject("ADODB.Recordset")
		'此段在比較資料庫與畫面的差別
		
		'當資料庫>畫面時,將多的從資料庫中刪除
		'response.write cint(request(maxAttach_no)) &">="& cint(request(filenum)) &"<BR>"
		'response.end
		IF cint(request(maxAttach_no)) >= cint(request(filenum)) then
			'response.write "當資料庫>=畫面時" & "<BR>"
			for i=1 to cint(request(filenum))
				'response.write "attach_no="&request(uploadfield&"_attach_no"&i)&"<Br>"
				'如果path沒有值時,則代表刪除該筆資料,否則用update的方式
				IF trim(request(uploadfield&i))<>empty then
					if trim(request(uploadfield & "_open_flag" & i))="on" _
					or trim(request(uploadfield & "_open_flag" & i))="Y" then
						open_flag = "Y"
					else
						open_flag = "N"
					end if
					uSQL = "Update exp_attach set Source='"& psource &"'" &_
					       ",attach_path='"& request(uploadfield&i) &"'" &_
					       ",doc_type='"& request(uploadfield & "_doc_type" & i) &"'" &_
					       ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" &_
					       ",attach_name='"& request(uploadfield & "_name" & i) &"'" &_
					       ",attach_size='"& request(uploadfield & "_size" & i) &"'" &_
					       ",source_name='"& request(uploadfield & "_source_name" & i) &"'" &_
					       ",open_flag='"& open_flag &"'" &_
					       ",attach_flag='U'" & _
					       ",tran_date=getdate()" &_
					       ",tran_scode='"&  session("scode") &"'" &_
					       " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "更新資料 >= Update exp_attach"& i &"=" & uSQL & "<br><br>"
					'response.end
					
					Conn.execute uSQL
				Elseif request(uploadfield&"_attach_sqlno"&i)<>empty then
					'dsql="Delete from exp_attach where attach_sqlno=" & request(uploadfield&"_attach_sqlno"&i)
					dsql = "update exp_attach set attach_flag='D'"
					dsql = dsql & " where attach_sqlno=" & request(uploadfield&"_attach_sqlno"&i)
					'response.write "沒有path刪除該筆資料"& i &"="& dsql &"<br><br>"
					'response.end
					Conn.Execute dsql
				End IF
			next
					'response.end
			Pfilenum=cint(request(filenum))+1
			for i=Pfilenum to cint(request(maxAttach_no))
				'dsql="Delete from exp_attach"
				dsql = "update exp_attach set attach_flag='D'"
				dsql = dsql & " where attach_sqlno='" & request(uploadfield&"_attach_sqlno"&i) &"'"
				'Response.write "當資料庫>畫面時delete exp_attach "& i &"="&dsql&"<br><br>"
				Conn.Execute dsql
			next	
			'response.end
		ElseIF cint(request(maxAttach_no)) < cint(request(filenum)) then
			'response.write "當資料庫<畫面時" & "<BR>"
			qsql = "Select max(attach_no)+1 as mattach_no from exp_attach"
			if prgid="exp21" or prgid="exp22" then
				qsql = qsql & " where exp_sqlno=" & patt_sqlno
			else
				qsql = qsql & " where seq=" & pseq &" and seq1='"& pseq1 &"'"
				qsql = qsql & " and step_grade="& pstep_grade
				qsql = qsql & " and att_sqlno="& patt_sqlno  
			end if
			'response.write qsql &"<BR>"
			'response.end
			rsf.open qsql,conn,1,1
			IF not rsf.eof then
				attachno=rsf("mattach_no")
			End IF
			for i=1 to cint(request(maxAttach_no))
				'如果path沒有值時,則代表刪除該筆資料,否則用update的方式
				IF trim(request(uploadfield&i))<>empty then
					if trim(request(uploadfield & "_open_flag" & i))="on" _
					or trim(request(uploadfield & "_open_flag" & i))="Y" then
						open_flag = "Y"
					else
						open_flag = "N"
					end if
					uSQL = "Update exp_attach set Source='"& psource &"'" &_
					       ",attach_path='"& request(uploadfield&i) &"'" &_
					       ",doc_type='"& request(uploadfield & "_doc_type" & i) &"'" &_
					       ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" &_
					       ",attach_name='"& request(uploadfield & "_name" & i) &"'" &_
					       ",attach_size='"& request(uploadfield & "_size" & i) &"'" &_
					       ",source_name='"& request(uploadfield & "_source_name" & i) &"'" &_
					       ",open_flag='"& open_flag &"'" &_
					       ",attach_flag='U'" & _
					       ",tran_date=getdate()" &_
					       ",tran_scode='"&  session("scode") &"'" &_
					       " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'Response.write "更新資料 < Update exp_attach"& i &"=" & uSQL & "<br><br>"
					Conn.execute uSQL
				Elseif request(uploadfield&"_attach_sqlno"&i)<>empty then
					dsql = "update exp_attach set attach_flag='D'"
					dsql = dsql & " where attach_no=" & i
					dsql = dsql & " and attach_sqlno=" & request(uploadfield&"_attach_sqlno"&i)
					'response.write "path沒有值時dsql"& i &"=" & dsql &"<br><br>"
					'response.end
					Conn.Execute dsql
				End IF
			next
			Pfilenum=cint(request(attach_cnt))+1
			'Pfilenum=cint(request(maxAttach_no))+1
			'response.write Pfilenum &"<BR>"
			'response.write request(sqlnum) &"<BR>"
			for i=Pfilenum to cint(request(sqlnum))
				IF trim(request(uploadfield&i))<>empty then
					if trim(request(uploadfield & "_open_flag" & i))="on" _
					or trim(request(uploadfield & "_open_flag" & i))="Y" then
						open_flag = "Y"
					else
						open_flag = "N"
					end if
					fsql = "insert into exp_attach (exp_sqlno,Seq,seq1,step_grade,case_no,att_sqlno" &_
						   ",Source,in_date,in_scode,Attach_no,attach_path,doc_type,attach_desc" &_
						   ",Attach_name,Attach_size,source_name,attach_flag,Mark,open_flag,tran_date,tran_scode) values (" &_
						   chknumzero(request("exp_sqlno")) &",'"& trim(pseq) &"','"& trim(pseq1) &"'," &_
					   chknumzero(pstep_grade) &","& chkcharnull(request("case_no")) &","
					fsql = fsql & "'"& patt_sqlno &"','"& psource &"',"
					fsql = fsql & "getdate(),'"& session("scode")&"','"& attachno &"','"& request(uploadfield&i) &"',"
					fsql = fsql & "'"& trim(request(uploadfield & "_doc_type" & i)) &"',"
					fsql = fsql & "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"'," &_
						   "'"& trim(request(uploadfield & "_size" & i)) &"','"& request(uploadfield & "_source_name" & i) &"'," &_
						   "'A','','"& open_flag &"',getdate(),'"& session("scode") &"')"
					'response.write "當資料庫>畫面時insert exp_attach"& i &"=" &fsql&"<br><br>"
					attachno = attachno + 1
					Conn.execute fsql
				End IF	
			next	
			RSf.close		
		End IF
	End Select	
	'response.end
End Function

'處理上傳圖檔的部份
Function upin_exp_attach_for_job(pseq,pseq1,pstep_grade,pjob_sqlno,prcom_sqlno)
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
        if trim(request(uploadfield & "_temp_doc" & i))<>empty then
            doc_type = trim(request(uploadfield & "_temp_doc" & i))
        else
            doc_type = trim(request(uploadfield & "_doc_type" & i))
        end if
		select case trim(request(uploadfield & "_dbflag" & i))
			case "A"
				'當上傳路徑不為空的 and attach_sqlno為空的,才需要新增
				IF trim(request(uploadfield&i))<>empty then
						IF trim(request(uploadfield & "_open_flag" & i))=empty then
							topen_flag="N"
						Else	
							topen_flag="Y"
						End IF
					   fsql = "insert into exp_attach (exp_sqlno,Seq,seq1,step_grade,case_no,att_sqlno,job_sqlno,Source"
					   fsql = fsql & ",in_date,in_scode,Attach_no,attach_path,attach_desc"
					   fsql = fsql & ",Attach_name,Attach_size,attach_flag,Mark,tran_date,tran_scode,doc_type,source_name,prcom_sqlno,open_flag"
					   fsql = fsql & ",in_no,apattach_sqlno"
					   If lcase(prgid)="exp3a2" then
							fsql = fsql & ",tran_datef,tran_scodef"
					   End IF
					   fSQL = fsql & ") values ("
					   fsql = fsql & "'"& trim(request(uploadfield & "_exp_sqlno" & i)) &"','"& trim(pseq) &"','"& trim(pseq1) &"',"
					   fsql = fsql & "'"& trim(pstep_grade) &"','"& trim(request(uploadfield & "_case_no" & i)) &"','"& trim(request(uploadfield & "_att_sqlno" & i)) &"','"& trim(pjob_sqlno) &"',"
					   fsql = fsql & "'"& uploadsource &"',getdate(),'"& session("scode")&"','"& trim(request(uploadfield & "_attach_no" & i)) &"','"& request(uploadfield&i) &"',"
					   fsql = fsql & "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"',"
					   fsql = fsql & "'"& trim(request(uploadfield & "_size" & i)) &"',"
					   fsql = fsql & "'A','',getdate(),'"& session("scode") & "','" & doc_type & "','" & trim(request(uploadfield & "_source_name" & i)) & "'," & trim(prcom_sqlno)&",'" & trim(topen_flag) & "'"
					   fsql = fsql & ",'"& request("in_no") &"'"
					   fsql = fsql & ",'"& trim(request(uploadfield & "_apattach_sqlno" & i)) &"'"
					   If lcase(prgid)="exp3a2" then
							fsql = fsql & ",getdate(),'"& session("scode") &"'"
					   End IF
					   fsql = fsql & ")"
						response.write "資料庫無資料新增exp_attach"& i &"<br>" &fsql&"<br><br>"
						attachno = attachno + 1
						'response.end
						Conn.execute fsql
				end if	
			case "U"
					'當attach_sqlno <> empty時 , 而且上傳的路徑又是空的時候,表示要刪除該筆資料,而非修改
					if request(uploadfield&"_attach_sqlno"&i) <> empty and trim(request(uploadfield&i)) = empty then
						call insert_log_table(conn,"D",prgid,"exp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
				
						'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
						if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
							dsql = "update exp_attach set attach_flag='D'"
							If lcase(prgid)="exp3a2" then
							 	dsql = dsql & ",tran_datef=getdate(),tran_scodef='"& session("scode") &"'"
							End IF
							dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
							'response.write "delete exp_attach "& i &"="&dsql&"<br><br>"
							Conn.Execute dsql
						else
							'不需要處理,表示原本db就沒有值
						end if
					else
						call insert_log_table(conn,"U",prgid,"exp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
						IF trim(request(uploadfield & "_open_flag" & i))=empty then
							topen_flag="N"
						Else	
							topen_flag="Y"
						End IF
						uSQL = "Update exp_attach set Source='"& uploadsource &"'"
						uSQL = uSQL & ",attach_path='"& request(uploadfield&i) &"'"
						uSQL = uSQL & ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" 
						uSQL = uSQL & ",attach_name='"& request(uploadfield & "_name" & i) &"'"
						uSQL = uSQL & ",attach_size='"& request(uploadfield & "_size" & i) &"'"
						uSQL = uSQL & ",source_name='"& request(uploadfield & "_source_name" & i) &"'"
						uSQL = uSQL & ",doc_type='"& doc_type &"'"
						uSQL = uSQL & ",attach_flag='U'"
						uSQL = uSQL & ",open_flag='"& topen_flag &"'"
						uSQL = uSQL & ",apattach_sqlno='"& request(uploadfield & "_apattach_sqlno" & i) &"'"						
						uSQL = uSQL & ",tran_date=getdate()"
						uSQL = uSQL & ",tran_scode='"&  session("scode") &"'"
						If lcase(prgid)="exp3a2" then
							uSQL = uSQL & ",tran_datef=getdate()"
							uSQL = uSQL & ",tran_scodef='"&  session("scode") &"'"
						End IF
						uSQL = uSQL & " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
						'response.write "更新資料 < Update exp_attach"& i &"=" & uSQL & "<br><br>"
						'response.end
						Conn.execute uSQL
					end if
			
			case "D"
				call insert_log_table(conn,"D",prgid,"exp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))

				'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
				if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
					dsql = "update exp_attach set attach_flag='D'"
					If lcase(prgid)="exp3a2" then
					 	fsql = fsql & ",tran_datef=getdate(),tran_scodef='"& session("scode") &"'"
					End IF
					dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "delete exp_attach "& i &"="&dsql&"<br><br>"
					Conn.Execute dsql
				else
					'不需要處理,表示原本db就沒有值
				end if
		end select	
	next
	

	'response.end
End Function

'將某特定table的某一筆資料寫入該table
'ptable：table name
'pkey_field：key 值欄位名稱，如有多個欄位請用；隔開
'pkey_value：與 pkey_field 相互配合，如有多個欄位請用；隔開
'ikey_field：特定欄位名稱，如有多個欄位請用；隔開
'ikey_value：與 key_field 相互配合，寫入資料，如有多個欄位請用；隔開
'ppkkey：不需寫入的欄位
function insert_table_data(pconn,ptable,pkey_field,pkey_value,ikey_field,ikey_value,ppkkey)
	dim tisql
	dim tfield_str
	dim ar_key_field
	dim ar_key_value
	dim wsql
	dim ti
	
	set tRS = Server.CreateObject("ADODB.Recordset")
	
	tfield_str = ""
	tfield_value = ""
	ar_ikey_field = split(ikey_field,";")
	ar_ikey_value = split(ikey_value,";")
	
	tisql = "select b.name from sysobjects a, syscolumns b "
	tisql = tisql & " where a.id = b.id  and a.name = '" & ptable & "' and a.xtype='U' "
	tisql = tisql & " order by b.colid "
	
	tRS.open tisql,pconn,1,1
	while not tRS.eof
		chkvalue = "N"
		if ubound(split(ppkkey,";"&tRS("name")&";"))>0 then
			
		else
			tfield_str = tfield_str & tRS("name") & ","
		
			select case tRS("name")
				case "tran_scode","tr_scode"
					tfield_value = tfield_value &"'"& session("scode") & "',"
					chkvalue = "Y"
				case "tran_date","tr_date"
					tfield_value = tfield_value & "getdate(),"
					chkvalue = "Y"
				case else
					if instr(1,ikey_field,";") <> 0 then
						for ti2 = 0 to ubound(ar_ikey_field)
							if ar_ikey_field(ti2)=tRS("name") then
								tfield_value = tfield_value & "'"& ar_Ikey_value(ti2) & "',"
								chkvalue = "Y"
							else
								'tfield_value = tfield_value & tRS("name") & ","
							end if
						next
					else
						tfield_value = tfield_value & tRS("name") & ","
					end if
			end select
			if chkvalue <> "Y" then
				tfield_value = tfield_value & tRS("name") & ","
			end if
		end if
		tRS.MoveNext
	wend
	tRS.close
	'response.write tfield_str &"<BR>"
	'response.write tfield_value &"<BR>"
	'response.end
	
	tfield_str = left(tfield_str,len(tfield_str) - 1)
	tfield_value = left(tfield_value,len(tfield_value) - 1)
	
	ar_key_field = ""
	ar_key_value = ""
	wsql = ""
	if instr(1,pkey_field,";") <> 0 then
		ar_key_field = split(pkey_field,";")
		ar_key_value = split(pkey_value,";")
		for ti = 0 to ubound(ar_key_field)
			wsql = wsql & " and " & ar_key_field(ti) & " = '" & ar_key_value(ti) & "' "
		next
	else
		wsql = " and " & pkey_field & " = '" & pkey_value & "' "
	end if
	
	tisql = "insert into " & ptable & "(" & tfield_str & ")"
	tisql = tisql & " select " & tfield_value
	tisql = tisql & " from " & ptable
	tisql = tisql & " where 1 = 1 "
	tisql = tisql & wsql
	'response.write tisql & "<br>"
	'response.end	
	pconn.execute tisql
	set tRS = nothing
end function


'處理上傳圖檔的部份_將文入附件入到各區所
Function upin_exp_attach_for_branch(pconn,pbranch,pseq,pseq1,pstep_grade,pbr_sqlno)
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
						fsql = "insert into exp_attach (exp_sqlno,Seq,seq1,step_grade,case_no,att_sqlno,br_sqlno,job_sqlno,Source" &_
					   ",in_date,in_scode,Attach_no,attach_path,attach_desc" &_
					   ",Attach_name,Attach_size,attach_flag,Mark,tran_date,tran_scode,doc_type,source_name,open_flag) values (" &_
					   "'"& trim(request(uploadfield & "_exp_sqlno" & i)) &"','"& trim(pseq) &"','"& trim(pseq1) &"'," &_
					   "'"& trim(pstep_grade) &"','"& trim(request(uploadfield & "_case_no" & i)) &"','"& trim(request(uploadfield & "_att_sqlno" & i)) &"','"& trim(pbr_sqlno) &"','"& trim(pjob_sqlno) &"'," &_
					   "'"& uploadsource &"',getdate(),'"& session("scode")&"','"& trim(request(uploadfield & "_attach_no" & i)) &"','"& attach_path &"'," &_
					   "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"'," &_
					   "'"& trim(request(uploadfield & "_size" & i)) &"'," &_
					   "'A','',getdate(),'"& session("scode") & "','" & trim(request(uploadfield & "_temp_doc" & i)) & "','" & trim(request(uploadfield & "_source_name" & i)) & "','" & trim(topen_flag) & "')"
						'response.write "資料庫無資料新增exp_attach"& i &"<br>" &fsql&"<br><br>"
						attachno = attachno + 1
						pConn.execute fsql
						if err.number=0  then
							Call copyfile_tobranch(trim(request("branch")),trim(request(uploadfield&i))) 
						End IF	
				end if	
			case "U"
					'當attach_sqlno <> empty時 , 而且上傳的路徑又是空的時候,表示要刪除該筆資料,而非修改
					if request(uploadfield&"_attach_sqlno"&i) <> empty and trim(request(uploadfield&i)) = empty then
						call insert_log_table(conn,"D",prgid,"exp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
				
						'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
						if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
							dsql = "update exp_attach set attach_flag='D'"
							dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
							'response.write "delete exp_attach "& i &"="&dsql&"<br><br>"
							pConn.Execute dsql
						else
							'不需要處理,表示原本db就沒有值
						end if
					else
						call insert_log_table(pconn,"U",prgid,"exp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
						IF trim(request(uploadfield & "_open_flag" & i))=empty then
							topen_flag="N"
						Else	
							topen_flag="Y"
						End IF
						attach_path=replace(trim(request(uploadfield&i)),"/temp/"&pbranch,"")
						
						uSQL = "Update exp_attach set Source='"& uploadsource &"'" &_
						       ",attach_path='"& attach_path &"'" &_
						       ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" &_
						       ",attach_name='"& request(uploadfield & "_name" & i) &"'" &_
						       ",attach_size='"& request(uploadfield & "_size" & i) &"'" &_
						       ",source_name='"& request(uploadfield & "_source_name" & i) &"'" &_
						       ",doc_type='"& request(uploadfield & "_temp_doc" & i) &"'" &_
						       ",attach_flag='U'" & _
						       ",open_flag='"& topen_flag &"'" & _
						       ",tran_date=getdate()" &_
						       ",tran_scode='"&  session("scode") &"'" &_
						       " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
						'response.write "更新資料 < Update exp_attach"& i &"=" & uSQL & "<br><br>"
						'response.end
						pConn.execute uSQL
						if err.number=0  then
							'response.write trim(request(uploadfield&i)) &"<BR>"
							'response.end
							Call copyfile_tobranch(trim(request("branch")),trim(request(uploadfield&i))) 
						End IF	
					end if
			
			case "D"
				call insert_log_table(pconn,"D",prgid,"exp_attach","attach_sqlno",request(uploadfield&"_attach_sqlno"&i))
				'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
				if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
					dsql = "update exp_attach set attach_flag='D'"
					dsql = dsql & " where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "delete exp_attach "& i &"="&dsql&"<br><br>"
					pConn.Execute dsql
				else
					'不需要處理,表示原本db就沒有值
				end if
		end select	
	next
End Function
%>
