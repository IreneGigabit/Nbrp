<script Language="vbScript">
'�ˬd�O�_���Q�`�Ĭ��H
'pjob_branch: �t�ΰϩҧO�B pdept:�����O PT�Apdept_area:I�i�f/E�X�f
'pctrl_type(�w�d): A�N�z�H,AP�ӽФH,C�Ȥ�
'pkeytype: cappl_name:�ץ�W��(��)�Beappl_name: �ץ�W��(�^)�Bap_cname:���q/�渹/�H�W��(��)�Bap_ename:���q/�渹/�H�W��(�^)�B
'          ap_crep:�N��H(��)�Bap_erep:�N��H(�^)
'pkeydata: ����r
'ptype: A:�ӽФH�s�W�AB:�O�_���Ĥ@���x�o
'function check_ctrl_Ckeydata(pjob_branch,pdept_area,pkeytype,pkeydata,ptype,pno)
function check_ctrl_Ckeydata(pjob_branch,pdept,pdept_area,pctrl_type,pseq,pseq1,pkeytype,pkeydata,ptype,pno)
    check_ctrl_Ckeydata = "Z"
    if trim(pkeydata)=empty then
    '    msgbox "����r�п�J"
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
                ans = msgbox("���ӽФH���Q�`�Ĭ��H�O�_�O����t�Τ� !!!",vbYesNo)
                if ans=6 then
                    reg.save_conflict_rec.value = "Y" '�J�ɥ�
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
                        check_ctrl_Ckeydata = "2"  '���O�Ĥ@��
                    else
                        check_ctrl_Ckeydata = "1"
                        ans = msgbox("���ץ�ӽФH���Q�`�Ĭ��H�O�_�O����t�Τ� !!!",vbYesNo)
                        if ans=6 then
                            if pno<>0 then
                                execute "reg.save_conflict_rec"&pno&".value = ""Y""" '�J�ɥ�
                            else
                                reg.save_conflict_rec.value = "Y" '�J�ɥ�
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
