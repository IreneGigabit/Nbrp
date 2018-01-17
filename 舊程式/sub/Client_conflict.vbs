<script Language="vbScript">
'檢查是否為利害衝突對象
'pjob_branch: 系統區所別、 pdept:部門別 PT，pdept_area:I進口/E出口
'pctrl_type(預留): A代理人,AP申請人,C客戶
'pkeytype: cappl_name:案件名稱(中)、eappl_name: 案件名稱(英)、ap_cname:公司/行號/人名稱(中)、ap_ename:公司/行號/人名稱(英)、
'          ap_crep:代表人(中)、ap_erep:代表人(英)
'pkeydata: 關鍵字
'ptype: A:申請人新增，B:是否為第一次官發
'function check_ctrl_Ckeydata(pjob_branch,pdept_area,pkeytype,pkeydata,ptype,pno)
function check_ctrl_Ckeydata(pjob_branch,pdept,pdept_area,pctrl_type,pseq,pseq1,pkeytype,pkeydata,ptype,pno)
    check_ctrl_Ckeydata = "Z"
    if trim(pkeydata)=empty then
    '    msgbox "關鍵字請輸入"
        exit function 
    end if

    SearchSql = "select b.detail_sqlno,b.conflict_sqlno,a.branch,a.dept,a.dept_area"
    SearchSql = SearchSql & " from conflict_main a"
    SearchSql = SearchSql & " inner join conflict_detail b on a.conflict_sqlno=b.conflict_sqlno and b.stat_code<>'D'"
    'SearchSql = SearchSql & " and b.keydata like '%"& ToUnicode(pkeydata) &"%'"
    SearchSql = SearchSql & " where a.ctrl_dates<='"& date() &"' and a.ctrl_datee>='"& date() &"'"
	url = "../xml/XmlGetSqlData_unicode_sidbs.asp?SearchSql=" & SearchSql &"&keydata="& ToUnicode(pkeydata)
	'window.open url
    Set xmldocs = CreateObject("Microsoft.XMLDOM")
	xmldocs.async = false
	xmldocs.validateOnParse = true
	If xmldocs.load(url) Then
		if xmldocs.selectSingleNode("//xhead/Found").text="Y" then
		    check_ctrl_Ckeydata = "A"
            if ptype="A" then
                ans = msgbox("此申請人為利害衝突對象是否記錄於系統中 !!!",vbYesNo)
                if ans=6 then
                    reg.save_conflict_rec.value = "Y" '入檔用
                end if
            elseif ptype="B" then
                SearchSql = "select * from conflict_rec"
                SearchSql = SearchSql & " where detail_sqlno="& xmldocs.selectSingleNode("//xhead/detail_sqlno").text
                SearchSql = SearchSql & " and syscode='"& "<%=session("Syscode")%>" &"' and branch='"& pjob_branch &"' and dept_area='"& pdept_area &"'"
                SearchSql = "select * from conflict_list"
                SearchSql = SearchSql & " where detail_sqlno="& xmldocs.selectSingleNode("//xhead/detail_sqlno").text 
                SearchSql = SearchSql & " and branch='"& pjob_branch &"' and dept='"& pdept&"'"
                SearchSql = SearchSql & " and dept_area='"& pdept_area &"' and stat_code<>'D' and seq='"& pseq &"' and seq1='"& pseq1 &"'"
	            url = "../xml/XmlGetSqlData_unicode_sidbs.asp?SearchSql=" & SearchSql
	            'window.open url
                Set xmldocs2 = CreateObject("Microsoft.XMLDOM")
	            xmldocs2.async = false
	            xmldocs2.validateOnParse = true
	            If xmldocs2.load(url) Then
		            if xmldocs2.selectSingleNode("//xhead/Found").text="Y" then
                        check_ctrl_Ckeydata = "2"  '表不是第一次
                    else
                        check_ctrl_Ckeydata = "1"
                        ans = msgbox("此案件申請人為利害衝突對象是否記錄於系統中 !!!",vbYesNo)
                        if ans=6 then
                            if pno<>0 then
                                execute "reg.save_conflict_rec"&pno&".value = ""Y""" '入檔用
                            else
                                reg.save_conflict_rec.value = "Y" '入檔用
                            end if
                        end if
                    end if
                end if
                set xmldocs2 = nothing 
            end if
	    end if
    end if
    set xmldocs = nothing
    
end function
</script>
