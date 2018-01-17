<%
function get_do_status_name(pdo_status)
    select case pdo_status
        case "NN" get_do_status_name = "尚未處理"
        case "NX" get_do_status_name = "退回營洽"
        case "DC" get_do_status_name = "程序確認中"
        case "SS" get_do_status_name = "主管簽核中"
        case "AA" get_do_status_name = "會計確認中"
        case "YY" get_do_status_name = "會計確認完成"
        case "YS" get_do_status_name = "確認完成(無需會計確認)"
        case "YB" get_do_status_name = "確認完成(無需處理)" '上月結算與最近規費加總為0，無需處理
    end select
end function
'---扣收入種類
function getDS_type(pconn,pType,pcho,pvalue,pwhere)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type = 'accDS_type'"
	if pwhere<>empty then
	    fsql = fsql & pwhere
	end if
	fsql = fsql & " order by sortfld"
	getDS_type = getcodeoption1(pconn,fsql,pType,pcho,"Y")
end function
'最近異動規費金額
function get_now_fees(pconn,pendtr_yy,pendtr_mm,pdept,pseq,pseq1,pcountry)
    dim arget_now_fees(5)
    Set rsf = Server.CreateObject("ADODB.RecordSet") 	
    nfees1 = 0
    nfees2 = 0
    nfees3 = 0
    nfees4 = 0
    nfees = 0
    ndate = pendtr_yy &"/"& pendtr_mm &"/1"
    'response.Write ndate & "<BR>"
    'response.End 
    ndate = dateadd("m",1,ndate)
    
    isql = "select isnull(sum(fees),0) as fees from acct_pin "
    isql = isql & " where branch='"& session("se_branch") &"' and dept='"& pdept &"'"
    isql = isql & " and seq="& pseq &" and seq1='"& pseq1 &"'"
    isql = isql & " and country='"& pcountry &"'"
    isql = isql & " and in_date>='"& ndate &"'"
    'if session("scode")="m983" then response.Write isql & "<BR>"
    'response.End
    rsf.Open isql,pconn,1,3
    while not rsf.EOF
        nfees1 = nfees1 + cdbl(rsf("fees"))
        rsf.MoveNext 
    wend
    rsf.Close 

    isql = "select * from acct_plus "
    isql = isql & " where branch='"& session("se_branch") &"' and dept='"& pdept &"'"
    isql = isql & " and seq="& pseq &" and seq1='"& pseq1 &"'"
    isql = isql & " and country='"& pcountry &"'"
    isql = isql & " and plus_date>='"& ndate &"'"
    if RSreg("ar_type")="t" then
        isql = isql & " and acc_code='21"& pdept &"0'"
    elseif RSreg("ar_type")="e" then
        isql = isql & " and acc_code='21"& pdept &"E'"
    end if
    'if session("scode")="m983" then response.Write isql & "<BR>"
    rsf.Open isql,pconn,1,3
    while not rsf.EOF
        if trim(rsf("acc_code1"))="01" then
            'if trim(rsf("dc_code"))="2" then '貸方
                nfees2 = nfees2 + cdbl(rsf("nt_money"))
            'else
            '    nfees4 = nfees4 + cdbl(rsf("nt_money"))
            'end if
        elseif trim(rsf("acc_code1"))="02" then
            'if trim(rsf("dc_code"))="2" then '貸方
                nfees3 = nfees3 + cdbl(rsf("nt_money"))
            'else
            '    nfees1 = nfees1 + cdbl(rsf("nt_money"))
            'end if
        elseif trim(rsf("acc_code1"))="99" then
            if trim(rsf("dc_code"))="2" then '貸方
                nfees2 = nfees2 + cdbl(rsf("nt_money"))
            else
                nfees4 = nfees4 + cdbl(rsf("nt_money"))
            end if
        end if
        rsf.MoveNext 
    wend
    rsf.Close 
    nfees = cdbl(nfees1) + cdbl(nfees2) - cdbl(nfees3) - cdbl(nfees4)

    arget_now_fees(0) = nfees1
    arget_now_fees(1) = nfees2
    arget_now_fees(2) = nfees3
    arget_now_fees(3) = nfees4
    arget_now_fees(4) = nfees
    get_now_fees = arget_now_fees
End Function
'未請款、已請款、未官發、已官發、未銷帳、已銷帳
function get_br_arfees(pconn,par_type,pseq,pseq1)
    dim arget_br_arfees(6)
    Set rsf = Server.CreateObject("ADODB.RecordSet")
    
    fnoar_fees = 0
    far_fees = 0
    fnogs_fees = 0
    if par_type="t" then
        isql = "select isnull(sum(fees),0)+isnull(sum(add_fees),0)-isnull(sum(ar_fees),0) as noar_fees,isnull(sum(ar_fees),0) as ar_fees"
        isql = isql & ",isnull(sum(fees),0)+isnull(sum(add_fees),0)-isnull(sum(gs_fees),0) as nogs_fees from case_dmp"
        isql = isql & " where seq="& pseq &" and seq1='"& pseq1 &"' and stat_code='YZ'"
        rsf.Open isql,conn,1,3
        if not rsf.EOF then
            fnoar_fees = rsf("noar_fees") '未請款 
            far_fees = rsf("ar_fees")     '已請款
            fnogs_fees = rsf("nogs_fees") '未官發
        end if
        rsf.Close
        ftr_money = 0
        isql = "select isnull(sum(tr_money),0) as tr_money"
        isql = isql & " from plus_temp"
        isql = isql & " where branch='"& session("se_branch") &"' and dept='"& session("dept") &"'"
        isql = isql & " and seq="& pseq &" and seq1='"& pseq1 &"' and mstat_flag='YY'"
        rsf.Open isql,connacc,1,3
        if not rsf.EOF then
            ftr_money = rsf("tr_money")  '已官發
        end if
        rsf.Close
        fnopin_fees = 0
        fpin_fees = 0
        isql = "select isnull(sum(fees),0)+isnull(sum(cfees),0)-isnull(sum(pin_fees),0) as nopin_fees,isnull(sum(pin_fees),0) as pin_fees"
        if session("dept")="P" then
            isql = isql & " from arp"
        elseif session("dept")="T" then
            isql = isql & " from art"
        end if
        isql = isql & " where branch='"& session("se_branch") & session("dept") &"'"
        isql = isql & " and country='T'"
        isql = isql & " and seq="& RSreg("seq") &" and seq1='"& RSreg("seq1") &"'"
        isql = isql & " and change<>'X'"
        rsf.Open isql,connacc,1,3
        if not rsf.EOF then
            fnopin_fees = rsf("nopin_fees") '未銷帳
            fpin_fees = rsf("pin_fees")     '已銷帳
        end if
        rsf.Close
        arget_br_arfees(0) = fnoar_fees
        arget_br_arfees(1) = far_fees
        arget_br_arfees(2) = fnogs_fees
        arget_br_arfees(3) = ftr_money
        arget_br_arfees(4) = fnopin_fees
        arget_br_arfees(5) = fpin_fees
    elseif par_type="e" then
        fnoar_fees = 0
        far_fees = 0
        fnogs_fees = 0
        isql = "select isnull(sum(a.fees),0)+isnull(sum(a.add_fees),0)-isnull(sum(a.ar_fees),0) as noar_fees,isnull(sum(a.ar_fees),0) as ar_fees"
        isql = isql & ",isnull(sum(a.fees),0)+isnull(sum(a.add_fees),0)-isnull(sum(a.gs_fees),0) as nogs_fees"
        'isql = isql & ",(select isnull(sum(fees),0) as tr_money from fees_exp where case_no=a.case_no) as tr_money"
        isql = isql &" from case_exp a"
        isql = isql & " where a.seq="& RSreg("seq") &" and a.seq1='"& RSreg("seq1") &"' and a.stat_code='YZ'"
        rsf.Open isql,conn,1,3
        if not rsf.EOF then
            fnoar_fees = rsf("noar_fees")  '未請款
            far_fees = rsf("ar_fees")      '已請款
            fnogs_fees = rsf("nogs_fees")  '未代收
        end if
        rsf.Close
        ftr_money = 0
        isql = "select isnull(sum(b.fees),0) as tr_money "
        isql = isql & " from case_exp a inner join fees_exp b on b. case_no=a.case_no"
        isql = isql & " where a.seq="& RSreg("seq") &" and a.seq1='"& RSreg("seq1") &"' and a.stat_code='YZ'"
        rsf.Open isql,conn,1,3
        if not rsf.EOF then
            ftr_money = rsf("tr_money")
        end if
        rsf.Close
        Set conndebit = Server.CreateObject("ADODB.Connection")
        conndebit.Open Session("debit")
        fnodn_nt_money = 0
        fdn_nt_money = 0
        isql = "select isnull(sum(dn_nt_money),0) as dn_nt_money from exch"
        isql = isql & " where dept='"& session("dept") &"' and seq="& RSreg("seq") &" and seq1='"& RSreg("seq1") &"'"
        isql = isql & " and appl_date is null and cancel_flag<>'Y'"
        rsf.Open isql,conndebit,1,3
        if not rsf.EOF then
            fnodn_nt_money = rsf("dn_nt_money")
        end if
        rsf.Close
        isql = "select isnull(sum(dn_nt_money),0) as dn_nt_money from exch"
        isql = isql & " where dept='"& session("dept") &"' and seq="& RSreg("seq") &" and seq1='"& RSreg("seq1") &"'"
        isql = isql & " and appl_date is not null and cancel_flag<>'Y'"
        rsf.Open isql,conndebit,1,3
        if not rsf.EOF then
            fdn_nt_money = rsf("dn_nt_money")
        end if
        rsf.Close
        conndebit.Close 
        set conndebit = nothing
        arget_br_arfees(0) = fnoar_fees
        arget_br_arfees(1) = far_fees
        arget_br_arfees(2) = fnogs_fees
        arget_br_arfees(3) = ftr_money
        arget_br_arfees(4) = fnodn_nt_money
        arget_br_arfees(5) = fdn_nt_money
    end if
    get_br_arfees = arget_br_arfees
End Function
'未請款、已請款尚未銷帳、尚未官發、已官發尚未入帳
function get_br_arfees2(pconn,pdept,par_type,pseq,pseq1)
    dim arget_br_arfees(4)
    Set rsf = Server.CreateObject("ADODB.RecordSet")
    
    if par_type="t" then
        Set fconnbr = Server.CreateObject("ADODB.Connection")
        fconnbr.Open Session("btbrtdb")
        fnoar_fees = 0
        isql = "select isnull(sum(fees),0)+isnull(sum(add_fees),0)-isnull(sum(ar_fees),0) as noar_fees"
        if pdept="P" then
            isql = isql & " from case_dmp"
        else
            isql = isql & " from case_dmt"
        end if
        isql = isql & " where seq="& pseq &" and seq1='"& pseq1 &"' and stat_code='YZ' and ar_code='N' "
        rsf.Open isql,fconnbr,1,3
        if not rsf.EOF then
            fnoar_fees = rsf("noar_fees") '未請款 
        end if
        rsf.Close
        fnogs_fees = 0
        isql = "select isnull(sum(fees+add_fees-gs_fees),0) as nogs_fees"
        if pdept="P" then
            isql = isql & " from case_dmp"
        else
            isql = isql & " from case_dmt"
        end if
        isql = isql & " where seq="& pseq &" and seq1='"& pseq1 &"' and stat_code='YZ'"
        isql = isql & " and case_date>='2008/1/1'"
        isql = isql & " having sum(fees+add_fees-gs_fees)<>0"
        rsf.Open isql,fconnbr,1,3
        if not rsf.EOF then
            fnogs_fees = rsf("nogs_fees") '提列規費
        end if
        rsf.Close
        fconnbr.Close 
        set fconnbr = nothing
        
        Set fconnacc = Server.CreateObject("ADODB.Connection")
        if session("syscode")="NACC" or session("syscode")="CACC" or session("syscode")="SACC" or session("syscode")="KACC" then
            fconnacc.Open session("acc")
        else
            fconnacc.Open session("sin09account")
        end if
        
        fgs_nopin_fees = 0
        'isql = "select isnull(sum(gs_fees),0) as gs_fees"
        'if pdept="P" then
        '    isql = isql & " from case_dmp"
        'else
        '    isql = isql & " from case_dmt"
        'end if
        'isql = isql & " where seq="& pseq &" and seq1='"& pseq1 &"' and stat_code='YZ' and gs_fees>0 and ar_fees=0"
        isql = "select isnull(sum(tr_money),0) as gs_fees"
        isql = isql & " from plus_temp "
        isql = isql & " where dept = '" & pdept & "' "
        isql = isql & " and seq="& pseq &" and seq1='"& pseq1 &"'"
        isql = isql & " and chk_type = 'N' and mstat_flag like 'Y%' "        
        
        rsf.Open isql,fconnacc,1,3
        if not rsf.EOF then
            fgs_nopin_fees = rsf("gs_fees") '已官發尚未入帳
        end if
        rsf.Close
        fnopin_fees = 0
        isql = "select isnull(sum(fees),0)+isnull(sum(cfees),0)-isnull(sum(pin_fees),0) as nopin_fees"
        if pdept="P" then
            isql = isql & " from arp"
        elseif pdept="T" then
            isql = isql & " from art"
        end if
        isql = isql & " where branch='"& session("se_branch") & pdept &"'"
        isql = isql & " and country='T'"
        isql = isql & " and seq="& RSreg("seq") &" and seq1='"& RSreg("seq1") &"'"
        isql = isql & " and change<>'X' and ar_mark not in ('D','M')"
        rsf.Open isql,fconnacc,1,3
        if not rsf.EOF then
            fnopin_fees = rsf("nopin_fees") '已請款尚未銷帳
        end if
        rsf.Close
        fconnacc.Close 
        set fconnacc = nothing
        
        arget_br_arfees(0) = fnoar_fees '未請款
        arget_br_arfees(1) = fgs_nopin_fees  '已官發尚未入帳
        arget_br_arfees(2) = fnogs_fees '提列規費
        arget_br_arfees(3) = fnopin_fees  '已請款尚未銷帳
    elseif par_type="e" then
        Set fconnbr = Server.CreateObject("ADODB.Connection")
        fconnbr.Open Session("btbrtdb")
        fnoar_fees = 0
        far_fees = 0
        if pdept="P" then
            isql = "select isnull(sum(a.fees),0)+isnull(sum(a.add_fees),0)-isnull(sum(a.ar_fees),0) as noar_fees"
            isql = isql &" from case_exp a"
        else
            isql = "select isnull(sum(a.tot_fees),0)+isnull(sum(a.add_fees),0)-isnull(sum(a.ar_fees),0) as noar_fees"
            isql = isql &" from case_ext a"
        end if
        isql = isql & " where a.seq="& RSreg("seq") &" and a.seq1='"& RSreg("seq1") &"' and a.stat_code='YZ' and ar_code='N'"
        rsf.Open isql,fconnbr,1,3
        if not rsf.EOF then
            fnoar_fees = rsf("noar_fees")  '未請款
        end if
        rsf.Close
        fnogs_fees = 0
        if pdept="P" then
            isql = "select isnull(sum(fees+add_fees-gs_fees),0) as nogs_fees"
            isql = isql &" from case_exp a"
            isql = isql & " where a.seq="& RSreg("seq") &" and a.seq1='"& RSreg("seq1") &"'"
        else
            isql = "select isnull(sum(a.tot_fees+a.add_fees-a.gs_fees),0) as nogs_fees"
            isql = isql &" from case_ext a"
            isql = isql & " where a.seq="& RSreg("seq") &" and a.seq1='"& RSreg("seq1") &"'"
        end if
        isql = isql & " and case_date>='2008/1/1'"
        if pdept="P" then
            isql = isql & " and a.stat_code='YZ'"
        else
            isql = isql & " and (a.stat_code = 'YZ' or a.stat_code like 'S%')"
        end if
        if pdept="P" then
            isql = isql & " having sum(fees+add_fees-gs_fees)<>0"
        else
            isql = isql & " having sum(tot_fees+add_fees-gs_fees)<>0"
        end if
        rsf.Open isql,fconnbr,1,3
        if not rsf.EOF then
            fnogs_fees = rsf("nogs_fees")  '未代收
        end if
        rsf.Close
   '     fgs_nopin_fees = 0
  '      isql = "select isnull(sum(gs_fees),0) as gs_fees"
 '       isql = isql & " from case_exp"
'        isql = isql & " where seq="& pseq &" and seq1='"& pseq1 &"' and stat_code='YZ' and gs_fees>0 and ar_fees=0"
    '    rsf.Open isql,fconnbr,1,3
   '     if not rsf.EOF then
  '          fgs_nopin_fees = rsf("gs_fees") '已官發尚未入帳
     '   end if
 '       rsf.Close
'        fconnbr.Close 
'        set fconnbr = nothing

        Set fconndebit = Server.CreateObject("ADODB.Connection")
        fconndebit.Open Session("debit")
        fnodn_nt_money = 0
        fdn_nt_money = 0
        isql = "select isnull(b.dis_money,0) as dis_money,isnull(b.dn_money,0) as dn_money,a.dn_rate,b.rate"
        isql = isql & " from exch_temp a inner join exch b on a.exch_no=b.exch_no"
        isql = isql & " where b.dept='"& pdept &"' and b.br_branch='"& session("se_branch") &"' and b.br_no="& RSreg("seq") &" and b.br_no1='"& RSreg("seq1") &"'"
        isql = isql & " and b.br_date is null and b.cancel_flag<>'Y'"
        'isql = isql & " and a.exch_no=b.exch_no and b.appl_date is null and b.cancel_flag<>'Y'"
        rsf.Open isql,fconndebit,1,3
        while not rsf.EOF   '已請款未結匯
            if cdbl(rsf("rate"))<>0 then
                dn_rate = rsf("rate")
            else
                dn_rate = rsf("dn_rate")
            end if
            if cdbl(rsf("dis_money"))=0 then
                fnodn_nt_money = fnodn_nt_money + (cdbl(rsf("dn_money")) * cdbl(dn_rate))
            else
                fnodn_nt_money = fnodn_nt_money + (cdbl(rsf("dis_money")) * cdbl(dn_rate))
            end if
            rsf.movenext
        wend
        rsf.Close
        fnodn_nt_money = fnodn_nt_money
        fconndebit.Close 
        set fconndebit = nothing

        Set fconnacc = Server.CreateObject("ADODB.Connection")
        if session("syscode")="NACC" or session("syscode")="CACC" or session("syscode")="SACC" or session("syscode")="KACC" then
            fconnacc.Open session("acc")
        else
            fconnacc.Open session("sin09account")
        end if
        fgs_nopin_fees = 0
        isql = "select isnull(sum(fees),0)+isnull(sum(cfees),0)-isnull(sum(pin_fees),0) as nopin_fees"
        if pdept="P" then
            isql = isql & " from arp"
        elseif pdept="T" then
            isql = isql & " from art"
        end if
        isql = isql & " where branch='"& session("se_branch") & pdept &"'"
        isql = isql & " and country<>'T'"
        isql = isql & " and seq="& RSreg("seq") &" and seq1='"& RSreg("seq1") &"'"
        isql = isql & " and change<>'X' and ar_mark not in ('D','M')"
        rsf.Open isql,fconnacc,1,3
        if not rsf.EOF then
            fgs_nopin_fees = rsf("nopin_fees") '已請款尚未銷帳
        end if
        rsf.Close
        fconnacc.Close 
        set fconnacc = nothing
        
        arget_br_arfees(0) = formatnumber(fnoar_fees,0)
        arget_br_arfees(1) = formatnumber(fgs_nopin_fees,0)
        arget_br_arfees(2) = formatnumber(fnogs_fees,0)
        arget_br_arfees(3) = formatnumber(fnodn_nt_money,0)
    end if
    get_br_arfees2 = arget_br_arfees
End Function
'已規費結餘轉收入
function get_acct_plusAA1(pconn,pdept,par_type,pseq,pseq1)
    Set rsf = Server.CreateObject("ADODB.RecordSet")
    
    get_acct_plusAA1 = "0"
'    isql = "select isnull(sum(nt_money),0) as nt_money from acct_plus where branch='"& session("se_branch") &"' and dept='"& pdept &"'"
'    isql = isql & " and seq='"& pseq &"' and seq1='"& pseq1 &"'"
'    if par_type="t" then
'        isql = isql & " and country='T' and country<>'z'"
'    elseif par_type="e" then
'        isql = isql & " and country<>'T'"
'    end if
'    isql = isql & " and acc_code1='02' and dc_code='1' and mark_code='AA1'"
    
    isql = "select isnull(sum(nt_money),0) as nt_money from acct_plus "
    isql = isql & " where branch='"& session("se_branch") &"' and dept='"& pdept &"'"
    isql = isql & " and seq="& pseq &" and seq1='"& pseq1 &"'"
    isql = isql & " and acc_code1='02'"
    if par_type="t" then
        isql = isql & " and acc_code='21"& pdept &"0'"
        isql = isql & " and country='T' and country<>'z'"
    elseif par_type="e" then
        isql = isql & " and acc_code='21"& pdept &"E'"
        isql = isql & " and country<>'T'"
    end if   
    'response.Write isql & "<BR>"
    rsf.open isql,pconn,1,3
    if not rsf.eof then
        get_acct_plusAA1 = rsf("nt_money")
    end if
    rsf.close
    get_acct_plusAA1 = formatnumber(get_acct_plusAA1,0)
    set rsf = nothing    
End Function
'顯示英文invoice
function show_edb_file(pconn,ar_no)
    Set rsf = Server.CreateObject("ADODB.RecordSet")
    
    isql = "select edb_file from arpmain "
    isql = isql & " where ar_no='" &ar_no& "' and edb_file is not null and edb_file<>'' "
    'response.Write isql & "<BR>"
    rsf.open isql,pconn,1,3
    if not rsf.eof then
        edb_file = rsf("edb_file")
		show_edb_file="<IMG border=0 src=""../images/annex.gif"" onclick=""window.open('"&session("brdbfile_path")&"/"&edb_file&"')"" style='cursor:pointer'>"
    end if
    rsf.close
    set rsf = nothing    
End Function
%>
