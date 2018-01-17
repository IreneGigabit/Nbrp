<%
'for 無共同log欄位table insert
'***新增arpmain_log
function insert_armain_log(pconn,ptblname,par_no,pbranch,pupd_flag)
   isql="insert into " & ptblname & "_log(ar_no,branch,ar_type,in_scode,in_date,ar_date,ar_scode,ar_company,"
   isql=isql & "ar_company1,invoice_mark,rec_mark,ar_id,acust_seq,att_sql,rec_chk,rec_chk1,cust_area,cust_seq,"
   isql=isql & "apsqlno,apcust_no,ar_zip,ar_addr1,ar_addr2,rec_type,print_chk,tax_chk,ctot_service,"
   isql=isql & "ctot_fees,tot_service,tot_fees,tot_money,tot_tax,tot_count,otot_service,otot_fees,otot_tax,dtot_service," 
   isql=isql & "dtot_fees,dtot_tax,ar_currency,ar_rate,tot_dnmoney,tot_rate_money,pre_money,pre_date,"
   isql=isql & "ar_chk,pay_date,pay_type,acchk_date,acc_chk,acc_name,exseq_chk,acc_date,rec_no,rec_no1,inv_no,inv_date,ar_status,"
   isql=isql & "new,tran_scode,tran_date,conf_scode,conf_date,mail_scode,mail_date,remark,mark_code,db_file,db_file_flag,db_file_scode,db_file_date,db_file_desc,db_curr,db_stat,opre_date,cust_paydate,"
   isql=isql & "del_scode,del_date,mark,trancase_sqlno,upd_flag) "
   isql=isql & " Select ar_no,branch,ar_type,in_scode,in_date,ar_date,ar_scode,ar_company,"
   isql=isql & "ar_company1,invoice_mark,rec_mark,ar_id,acust_seq,att_sql,rec_chk,rec_chk1,cust_area,cust_seq,"
   isql=isql & "apsqlno,apcust_no,ar_zip,ar_addr1,ar_addr2,rec_type,print_chk,tax_chk,ctot_service,"
   isql=isql & "ctot_fees,tot_service,tot_fees,tot_money,tot_tax,tot_count,otot_service,otot_fees,otot_tax,dtot_service," 
   isql=isql & "dtot_fees,dtot_tax,ar_currency,ar_rate,tot_dnmoney,tot_rate_money,pre_money,pre_date,"
   isql=isql & "ar_chk,pay_date,pay_type,acchk_date,acc_chk,acc_name,exseq_chk,acc_date,rec_no,rec_no1,inv_no,inv_date,ar_status,"
   isql=isql & "new,tran_scode,tran_date,conf_scode,conf_date,mail_scode,mail_date,remark,mark_code,db_file,db_file_flag,db_file_scode,db_file_date,db_file_desc,db_curr,db_stat,opre_date,cust_paydate,"
   isql=isql & "'" & session("se_scode") & "',getdate(),mark,'"&artran_sqlno&"','"&pupd_flag&"' "
   isql=isql & " from " & ptblname
   isql=isql & " where ar_no='"&par_no&"' and branch='"& pbranch &"'"
   'response.write "insert " & ptblname & "_log="&isql & "<BR><BR>"
   'Response.End
   pconn.Execute(isql)
end function

'入請款單明細檔log,artitem_log
function insert_aritem_log(pconn,ptblname,par_no,pbranch,pupd_flag,ptran_sqlno)
    isql="insert into " & ptblname & "_log(trancase_sqlno,upd_flag,item_sqlno,ar_no,branch,case_no,seq,seq1,country,arcase," 
    isql=isql & "item_count,cservice,cfees,aservice,afees,rservice,rfees,rar_money,rtax_out,rtr_money,tr_dept,orservice,"
    isql=isql & "orfees,ortax_out,drservice,drfees,drtax_out,ar_remark,ar_code,mark) "
    isql=isql & " select " & ptran_sqlno & ",'" & pupd_flag & "',item_sqlno,ar_no,branch,case_no,seq,seq1,country,arcase,"
    isql=isql & "item_count,cservice,cfees,aservice,afees,rservice,rfees,rar_money,rtax_out,rtr_money,tr_dept,orservice,"
    isql=isql & "orfees,ortax_out,drservice,drfees,drtax_out,ar_remark,ar_code,mark from " & ptblname
    isql=isql & " where ar_no='" & par_no & "' and branch='" & pbranch & "'"
    Response.Write "insert " & ptblname & "_log=" & isql & "<br>"
    pconn.execute(isql)  
end function
'入請款單明細檔log,arpitem1_log
function insert_aritem1_log(pconn,ptblname,par_no,pbranch,pupd_flag,ptran_sqlno)
    isql="insert into " & ptblname & "_log(trancase_sqlno,upd_flag,item1_sqlno,ar_no,branch,case_no,item_case," 
    isql=isql & "item_sql,item_service,item_fees,item_money,item_remark,mark) "
    isql=isql & " select " & ptran_sqlno & ",'" & pupd_flag & "',item1_sqlno,ar_no,branch,case_no,item_case,"
    isql=isql & "item_sql,item_service,item_fees,item_money,item_remark,mark from " & ptblname
    isql=isql & " where ar_no='" & par_no & "' and branch='" & pbranch & "'"
    Response.Write "insert " & ptblname & "_log=" & isql & "<br>"
    pconn.execute(isql)  
end function

'---insert chgarmain_log
'pfldname:table欄位名稱,pfldcname:欄位對應中文名稱
function insert_chgarmain_log(pconn,prgid,pbranch,par_type,pseq_area,par_no,pfldname,pfldcname,povalue,pnvalue)
    tsql="insert into chgarmain_log(branch,ar_type,seq_area,ar_no,fldname,fldcname,ovalue,nvalue,tran_date,tran_scode,prgid) values ("
    tsql=tsql & "'" & pbranch & "','" & par_type & "','" & pseq_area & "','" & par_no & "','" & pfldname & "','" & pfldcname & "'"
    tsql=tsql & ",'" & povalue & "','" & pnvalue & "',getdate(),'" & session("se_scode") & "','" & prgid & "')"
    'response.write "insert-chgarmain_log=" & tsql & "<br>"
    'response.end
    pconn.execute(tsql)
end function

%>