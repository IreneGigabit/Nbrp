<script Language="vbScript">
'�ˬd�O�_�i�߮צ���
'pjob_branch: �t�ΰϩҧO
'pkeytype: cappl_name:�ץ�W��(��)�Beappl_name: �ץ�W��(�^)�Bap_cname:���q/�渹/�H�W��(��)�Bap_ename:���q/�渹/�H�W��(�^)�B
'          ap_crep:�N��H(��)�Bap_erep:�N��H(�^)
'pkeydata: ����r
function check_ctrl_keydata(pjob_branch,pkeytype,pkeydata)
    check_ctrl_keydata = "Z"
    'if trim(pkeydata)=empty then
    '    msgbox "����r�п�J"
    '    exit function 
    'end if

    SearchSql = "select a.*,(select code_name from cust_code where code_type='keytype' and cust_code=b.keytype) as keytypenm"
    SearchSql = SearchSql & ",(select branchname from sysctrl.dbo.branch_code where branch=a.in_branch) as in_branchnm"
	SearchSql = SearchSql & " from query_main a inner join query_detail b on a.query_sqlno=b.query_sqlno"
	SearchSql = SearchSql & " and b.keytype='"& pkeytype &"' and b.keydata='"& ToUnicode(pkeydata) &"'"
	'SearchSql = SearchSql & " where a.in_branch<>'"& pjob_branch &"'"	'2016/12/23�h��ק�A�]�ũ��i���d�ӹ�H�b�o�d���]���ˬd�A����Ȫ��p�A�o�d��������|���d�ӹ�H�ץ�
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
		'    check_ctrl_keydata = "A"  '���i�߮צ���
		    'msgbox xmldocs.selectSingleNode("//xhead/keytypenm").text & "������N�z�d�ӹ�H���i�s�W !!!"
		    check_ctrl_keydata = "B"  '����
		    tyy = year(xmldocs.selectSingleNode("//xhead/ctrl_dates").text)
		    tmm = month(xmldocs.selectSingleNode("//xhead/ctrl_dates").text)
		    tdd = day(xmldocs.selectSingleNode("//xhead/ctrl_dates").text)
		    tyye = year(xmldocs.selectSingleNode("//xhead/ctrl_datee").text)
		    tmme = month(xmldocs.selectSingleNode("//xhead/ctrl_datee").text)
		    tdde = day(xmldocs.selectSingleNode("//xhead/ctrl_datee").text)
		    tmsg = "���ץӽФH�]�e�����^�ץ��"&tyy&"�~"&tmm&"��"&tdd&"��_�C�ަ�"&tyye&"�~"&tmme&"��"&tdde&"��A"
		    tmsg = tmsg & "�Ա��Ц�����N�z�d�ӵ��G�d�߿�J�y�����u"& xmldocs.selectSingleNode("//xhead/query_sqlno").text &"�v�C"
            tmsg = tmsg & chr(10)&chr(13) &"���J�׽Х���o�u"& xmldocs.selectSingleNode("//xhead/in_branchnm").text &"�v��줹�\�C"
            tmsg = tmsg & chr(10)&chr(13)&chr(10)&chr(13) &"���u�O�v��^�e���ק�A���u�_�v���~�����@�~ !!!"
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
    '    msgbox "����r�п�J"
    '    exit function 
    'end if

    SearchSql = "select a.*,(select code_name from cust_code where code_type='keytype' and cust_code=b.keytype) as keytypenm"
	SearchSql = SearchSql & " from query_main a inner join query_detail b on a.query_sqlno=b.query_sqlno"
	SearchSql = SearchSql & " and b.keytype='"& pkeytype &"' and b.keydata='"& ToUnicode(pkeydata) &"'"
	'SearchSql = SearchSql & " where a.in_branch<>'"& pjob_branch &"'"	'2016/12/23�h��ק�A�]�ũ��i���d�ӹ�H�b�o�d���]���ˬd�A����Ȫ��p�A�o�d��������|���d�ӹ�H�ץ�
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
		    check_ctrl_keydata2 = "A"  '���i�߮צ���
		    msgbox xmldocs.selectSingleNode("//xhead/keytypenm").text & "������N�z�d�ӹ�H���i�s�W !!!"
	    end if
    end if
    set xmldocs = nothing
    
end function
</script>
