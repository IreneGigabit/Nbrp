<!--#Include file="../sub/server_conn_unicode.vbs" -->
<% 
'利害衝突對象資料檢查及入檔
Set connsi = Server.CreateObject("ADODB.Connection")
connsi.Open session("sidbsdb")

'---傳入變數說明
'pbranch:區所別 NCSKTL， pdept:部門別 PT， pdept_area:I進口/E出口
'pfallseq:完整案件編號,ex:PI-12345-AM
'國外所 pagent_no, pagent_no1: 傳入本次立案案件的代理人號 或 新增的代理人號
'國外所 pap_area, pap_no, pap_no1: 傳入本次立案案件的申請人號 或 新增的申請人號
'區所 papsqlno, papcust_no: 傳入本次立案案件的申請人號 或 新增的申請人號

'利害衝突對象資料檢查
'---pkind: A:表客戶/申請人新增  B:表官發
'---pword: 傳入關鍵字
function check_conflict_data(pkind,pword,pbranch,pdept,pdept_area,pseq,pseq1,pfallseq,pfin_date,papply_date,pctrl_type,pap_area,pap_no,pap_no1,pagent_no,pagent_no1,papsqlno,papcust_no)
    check_conflict_data = ""
    if trim(pword)=empty then
        exit function 
    end if
    
    Set rsf = Server.CreateObject("ADODB.Recordset")
    Set rsf2 = Server.CreateObject("ADODB.Recordset")
    Set rsf3 = Server.CreateObject("ADODB.Recordset")
    isql = "select b.detail_sqlno,b.conflict_sqlno,a.branch,a.dept,a.dept_area"
    isql = isql & " from conflict_main a"
    isql = isql & " inner join conflict_detail b on a.conflict_sqlno=b.conflict_sqlno and b.stat_code<>'D'"
    isql = isql & " and b.keydata like '%"& pword &"%'"
    'isql = isql & " where a.ctrl_type='"& pctrl_type &"' and a.branch='"& pbranch &"' and a.dept='"& pdept &"'"
    'isql = isql & " and a.dept_area='"& pdept_area &"'"
'    if pkind="B" then
'        if pctrl_type="A" then
'            isql = isql & " and a.agent_no='"& pagent_no &"' and a.agent_no1='"& pagent_no1 &"'"
'        elseif pctrl_type="AP" and pbranch="T" then
'            isql = isql & " and a.ap_area='"& pap_area &"' and a.ap_no='"& pap_no &"' and a.ap_no1='"& pap_no1 &"'"
'        elseif pctrl_type="C" then
'            isql = isql & " and a.cust_area='"& pcust_area &"' and a.cust_seq='"& pcust_seq &"'"
'        end if
'    end if
    isql = isql & " and a.ctrl_dates<='"& date() &"' and a.ctrl_datee>='"& date() &"'"
'    isql = isql & " group by b.conflict_sqlno,b.detail_sqlno"
    'response.write isql & "<BR>"
    'response.end
    rsf.Open isql,connsi,1,3
    while not rsf.EOF
        'response.write "cnt="& rsf.RecordCount &"<BR>"
        if rsf.RecordCount>0 then
            '若已入conflict_rec不再入
            if pkind="A" then
                isql = "select * from conflict_rec"
                isql = isql & " where detail_sqlno="& rsf("detail_sqlno") &" and branch='"& pbranch &"' and ctrl_type='"& pctrl_type &"'"
                if pctrl_type="A" then
                    isql = isql & " and agent_no='"& pagent_no &"' and agent_no1='"& pagent_no1 &"'"
                elseif pctrl_type="AP" and pbranch="T" then
                    isql = isql & " and ap_area='"& pap_area &"' and ap_no='"& pap_no &"' and ap_no1='"& pap_no1 &"'"
                elseif pctrl_type="C" then
                    isql = isql & " and cust_area='"& pcust_area &"' and cust_seq='"& pcust_seq &"'"
                else
                    isql = isql & " and apcust_no='"& papcust_no &"'"
                end if
                'response.write isql & "<BR>"
                rsf2.Open isql,connsi,1,3
                if not rsf2.EOF then
                    check_conflict_data = "N"  '表不是第一次
                else
                    call insert_conflict_rec(session("se_branch"),session("dept"),pdept_area,pctrl_type,pap_area,pap_no,pap_no1,pagent_no,pagent_no1,papsqlno,papcust_no,rsf("conflict_sqlno"),rsf("detail_sqlno"))
                    if err.number = 0 then
                        check_conflict_data = "Y"
                    end if
                end if
                rsf2.Close 
            end if
            '若已入conflict_list不再入
            if pkind="B" then
                isql = "select * from conflict_list"
                isql = isql & " where detail_sqlno="& rsf("detail_sqlno") &" and branch='"& pbranch &"' and dept='"& pdept&"'"
                isql = isql & " and dept_area='"& pdept_area &"' and stat_code<>'D' and seq='"& pseq &"' and seq1='"& pseq1 &"'"
                'response.write isql & "<BR>"
                rsf2.Open isql,connsi,1,3
                if not rsf2.EOF then
                    nowlist_sqlno = rsf("list_sqlno")
                    check_conflict_data = "N"  '表不是第一次
                else
                    call insert_conflict_list(session("se_branch"),session("dept"),pdept_area,pseq,pseq1,pfallseq,pfin_date,papply_date,pap_area,pap_no,pap_no1,pagent_no,pagent_no1,papsqlno,papcust_no,rsf("conflict_sqlno"),rsf("detail_sqlno"))
                    if err.number = 0 then
                        check_conflict_data = "Y"
                        isql = "select list_sqlno from conflict_list"
                        isql = isql & " where detail_sqlno="& rsf("detail_sqlno") 
                        isql = isql & " order by list_sqlno desc"
                        rsf3.Open isql,connsi,1,3
                        if not rsf3.eof then
                            nowlist_sqlno = rsf3("list_sqlno")
                        end if
                        rsf3.Close 
                    end if
                end if
                rsf2.Close
                '檢查是不是該關鍵字的第一次發文
                isql = "select top 1 * from conflict_list"
                isql = isql & " where detail_sqlno="& rsf("detail_sqlno") 
                isql = isql & " order by list_sqlno"
                'response.write isql & "<BR>"
                rsf3.Open isql,connsi,1,3
                if not rsf3.eof then
                    if nowlist_sqlno<>rsf3("list_sqlno") then
                        'response.write "Y2" &"<BR>"
                        check_conflict_data = "N" '非第一次發文
                    end if
                end if
                rsf3.Close 
                'response.write "check_conflict_data=" & check_conflict_data &"<BR>"
            end if
        end if
        rsf.MoveNext 
    wend
    rsf.Close
end function
'利害衝突對象資料入檔
function insert_conflict_list(pbranch,pdept,pdept_area,pseq,pseq1,pfallseq,pfin_date,papply_date,pap_area,pap_no,pap_no1,pagent_no,pagent_no1,papsqlno,papcust_no,pconflict_sqlno,pdetail_sqlno)
    usql = "insert into conflict_list(conflict_sqlno,detail_sqlno,in_scode,in_date,syscode,in_prgid,branch,dept,dept_area,seq,seq1,fallseq"
    usql = usql &",fin_date,apply_date,ap_area,ap_no,ap_no1,agent_no,agent_no1,apsqlno,apcust_no,tran_date,tran_scode) "
    usql = usql &" values("& pconflict_sqlno &","& pdetail_sqlno &","& chkempty_unicode(session("scode")) &",getdate()"
    usql = usql &","& chkempty_unicode(session("syscode")) &","& chkempty_unicode(prgid) &","& chkempty_unicode(session("se_branch")) 
    usql = usql &","& chkempty_unicode(session("dept"))
    usql = usql &","& chkempty_unicode(pdept_area) &","& chknumzero(pseq) &","& chkempty_unicode(pseq1) &","& chkempty_unicode(pfallseq)
    usql = usql &","& chkdatenull(pfin_date) &","& chkdatenull(papply_date)
    usql = usql &","& chkempty_unicode(pap_area) &","& chkempty_unicode(pap_no) &","& chkempty_unicode(pap_no1)
    usql = usql &","& chkempty_unicode(pagent_no) &","& chkempty_unicode(pagent_no1) 
    usql = usql &","& chknumzero(papsqlno) &","& chkempty_unicode(papcust_no) 
    usql = usql &",getdate(),"& chkempty_unicode(session("scode"))
    usql = usql & ")"
    'response.write usql & "<BR>"
    'response.end
    connsi.Execute usql
    
end function

'若申請人新增時，該申請人為衝突對象需入conflict_rec
function insert_conflict_rec(pbranch,pdept,pdept_area,pctrl_type,pap_area,pap_no,pap_no1,pagent_no,pagent_no1,papsqlno,papcust_no,pconflict_sqlno,pdetail_sqlno)
    usql = "insert into conflict_rec(conflict_sqlno,detail_sqlno,in_scode,in_date,syscode,in_prgid,ctrl_type,branch,dept,dept_area,ap_area,ap_no,ap_no1"
    usql = usql & ",agent_no,agent_no1,apsqlno,apcust_no,tran_date,tran_scode)"
    usql = usql &" values("& pconflict_sqlno &","& pdetail_sqlno &","& chkempty_unicode(session("scode")) &",getdate()"
    usql = usql &","& chkempty_unicode(session("syscode")) &","& chkempty_unicode(prgid) &","& chkempty_unicode(pctrl_type) 
    usql = usql &","& chkempty_unicode(session("se_branch")) &","& chkempty_unicode(session("dept")) &","& chkempty_unicode(pdept_area)
    if pctrl_type="AP" then
        usql = usql &","& chkempty_unicode(pap_area) &","& chkempty_unicode(pap_no) &","& chkempty_unicode(pap_no1)
    else
        usql = usql &",'','',''"
    end if
    if pctrl_type="P" then
        usql = usql &","& chkempty_unicode(pagent_no) &","& chkempty_unicode(pagent_no1) 
    else
        usql = usql &",'',''"
    end if
    if pctrl_type="AP" and (pbranch="N" or pbranch="C" or pbranch="S" or pbranch="K") then
        usql = usql &","& chknumzero(papsqlno) &","& chkempty_unicode(papcust_no) 
    else
        usql = usql &",0,''"
    end if
    usql = usql &",getdate(),"& chkempty_unicode(session("scode"))
    usql = usql & ")"
    'response.write usql & "<BR>"
    'response.end
    connsi.Execute usql
end function
%>
