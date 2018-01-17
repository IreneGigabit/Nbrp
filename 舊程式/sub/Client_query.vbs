<script Language="vbScript">
'檢查是否可立案收文
'pjob_branch: 系統區所別
'pkeytype: cappl_name:案件名稱(中)、eappl_name: 案件名稱(英)、ap_cname:公司/行號/人名稱(中)、ap_ename:公司/行號/人名稱(英)、
'          ap_crep:代表人(中)、ap_erep:代表人(英)
'pkeydata: 關鍵字
function check_ctrl_keydata(pjob_branch,pkeytype,pkeydata)
    check_ctrl_keydata = "Z"
    'if trim(pkeydata)=empty then
    '    msgbox "關鍵字請輸入"
    '    exit function 
    'end if

    SearchSql = "select a.*,(select code_name from cust_code where code_type='keytype' and cust_code=b.keytype) as keytypenm"
    SearchSql = SearchSql & ",(select branchname from sysctrl.dbo.branch_code where branch=a.in_branch) as in_branchnm"
	SearchSql = SearchSql & " from query_main a inner join query_detail b on a.query_sqlno=b.query_sqlno"
	SearchSql = SearchSql & " and b.keytype='"& pkeytype &"' and b.keydata='"& ToUnicode(pkeydata) &"'"
	'SearchSql = SearchSql & " where a.in_branch<>'"& pjob_branch &"'"	'2016/12/23柳月修改，因嘉明告知查照對象在發查單位也需檢查，雖實務狀況，發查單位應不會接查照對象案件
	SearchSql = SearchSql & " where 1=1 "
	SearchSql = SearchSql & " and a.query_stat='YZ' and stat_code<>'N'"
	SearchSql = SearchSql & " and a.ctrl_dates<='"& date() &"' and a.ctrl_datee>='"& date() &"'"
	url = "../xml/XmlGetSqlData_unicode_sidbs.asp?SearchSql=" & SearchSql
	'window.open url
    Set xmldocs = CreateObject("Microsoft.XMLDOM")
	xmldocs.async = false
	xmldocs.validateOnParse = true
	If xmldocs.load(url) Then
		if xmldocs.selectSingleNode("//xhead/Found").text="Y" then
		'    check_ctrl_keydata = "A"  '不可立案收文
		    'msgbox xmldocs.selectSingleNode("//xhead/keytypenm").text & "為雙邊代理查照對象不可新增 !!!"
		    check_ctrl_keydata = "B"  '提示
		    tyy = year(xmldocs.selectSingleNode("//xhead/ctrl_dates").text)
		    tmm = month(xmldocs.selectSingleNode("//xhead/ctrl_dates").text)
		    tdd = day(xmldocs.selectSingleNode("//xhead/ctrl_dates").text)
		    tyye = year(xmldocs.selectSingleNode("//xhead/ctrl_datee").text)
		    tmme = month(xmldocs.selectSingleNode("//xhead/ctrl_datee").text)
		    tdde = day(xmldocs.selectSingleNode("//xhead/ctrl_datee").text)
		    tmsg = "本案申請人因前有爭訟案件自"&tyy&"年"&tmm&"月"&tdd&"日起列管至"&tyye&"年"&tmme&"月"&tdde&"日，"
		    tmsg = tmsg & "詳情請至雙邊代理查照結果查詢輸入流水號「"& xmldocs.selectSingleNode("//xhead/query_sqlno").text &"」。"
            tmsg = tmsg & chr(10)&chr(13) &"欲入案請先獲得「"& xmldocs.selectSingleNode("//xhead/in_branchnm").text &"」單位允許。"
            tmsg = tmsg & chr(10)&chr(13)&chr(10)&chr(13) &"按「是」表回畫面修改，按「否」表繼續執行作業 !!!"
            ans = msgbox(tmsg,vbYesNo)
            if ans=6 then
                check_ctrl_keydata = "A"
            end if
	    end if
    end if
    set xmldocs = nothing
    
end function

function check_ctrl_keydata2(pjob_branch,pkeytype,pkeydata)
    check_ctrl_keydata2 = "Z"
    'if trim(pkeydata)=empty then
    '    msgbox "關鍵字請輸入"
    '    exit function 
    'end if

    SearchSql = "select a.*,(select code_name from cust_code where code_type='keytype' and cust_code=b.keytype) as keytypenm"
	SearchSql = SearchSql & " from query_main a inner join query_detail b on a.query_sqlno=b.query_sqlno"
	SearchSql = SearchSql & " and b.keytype='"& pkeytype &"' and b.keydata='"& ToUnicode(pkeydata) &"'"
	'SearchSql = SearchSql & " where a.in_branch<>'"& pjob_branch &"'"	'2016/12/23柳月修改，因嘉明告知查照對象在發查單位也需檢查，雖實務狀況，發查單位應不會接查照對象案件
	SearchSql = SearchSql & " where 1=1 "
	SearchSql = SearchSql & " and a.query_stat='YZ' and stat_code<>'N'"
	SearchSql = SearchSql & " and a.ctrl_dates<='"& date() &"' and a.ctrl_datee>='"& date() &"'"
	url = "../xml/XmlGetSqlData_unicode_sidbs.asp?SearchSql=" & SearchSql
	'window.open url
    Set xmldocs = CreateObject("Microsoft.XMLDOM")
	xmldocs.async = false
	xmldocs.validateOnParse = true
	If xmldocs.load(url) Then
		if xmldocs.selectSingleNode("//xhead/Found").text="Y" then
		    check_ctrl_keydata2 = "A"  '不可立案收文
		    msgbox xmldocs.selectSingleNode("//xhead/keytypenm").text & "為雙邊代理查照對象不可新增 !!!"
	    end if
    end if
    set xmldocs = nothing
    
end function
</script>
