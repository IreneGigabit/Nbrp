<%
'---cust_code
function getcust_code(pwh1,pwh2,psort,pType,pcho)
	dim isql
	isql = "select cust_code,code_name from cust_code"
	isql = isql & " where code_type='"& pwh1 &"'"
	isql = isql & pwh2
	isql = isql & " order by "& psort &",cust_code"
	'getcust_code = isql
	getcust_code = getcodeoption(conn,isql,pType,pcho)
end function
function getcust_code1(pwh1,pwh2,psort,pType,pcho,pvalue)
	dim isql
	isql = "select cust_code,code_name from cust_code"
	isql = isql & " where code_type='"& pwh1 &"'"
	isql = isql & pwh2
	isql = isql & " order by "& psort &",cust_code"
	'getcust_code = isql
	getcust_code1 = getcodeoption1(conn,isql,pType,pcho,pvalue)
end function
'---radio getcust_code_mul(inputtype,pfldname,�Ѽ�1,�Ѽ�2,�ƧǤ覡,true��ܥN�X,disabled,checked��value,Y�n���n����)
function getcust_code_mul(inputtype,pfldname,pwh1,pwh2,psort,pType,pdisabled,pvalue,pbr)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type='"& pwh1 &"'"
	fsql = fsql & pwh2
	fsql = fsql & " order by convert(int,"& psort &"),cust_code"
	'getcust_code_mul = fsql
	getcust_code_mul = getcodeoption_mul(conn,fsql,inputtype,pfldname,pType,pdisabled,pvalue,pbr)
end function
function getcust_code_mul1(inputtype,pfldname,pwh1,pwh2,psort,pType,pdisabled,pvalue,pbr,ptabindex,ponclick)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type='"& pwh1 &"'"
	fsql = fsql & pwh2
	fsql = fsql & " order by convert(int,"& psort &"),cust_code"
	'response.write fsql & "<BR>"
	getcust_code_mul1 = getcodeoption_mul2(conn,fsql,inputtype,pfldname,pType,pdisabled,pvalue,pbr,ptabindex,ponclick)
end function
'�_�ʽдں���
function getag_flag(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type='peag_flag'"
	fsql = fsql & " order by sortfld,cust_code"
	getag_flag = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'�_�ʽдڳ���Y�n�D
function getco_type(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type='peco_type'"
	fsql = fsql & " order by sortfld,cust_code"
	getco_type = showselect5(conn,fsql,pType,pcho,pvalue)
end function

'---�ӽФH/�M�Q�v�H����
function getap_level(pfld,pfldtype,pType,pdisabled,pvalue,pbr,ptabindex)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type='"& session("dept") &"Eap_level'"
	fsql = fsql & " order by sortfld,cust_code"
	getap_level = getcodeoption_mul2(conn,fsql,pfldtype,pfld,pType,pdisabled,pvalue,pbr,ptabindex,"ap_level_onclick")
end function
'---�e�U�_�ʸt�qú�~�O
function getpay_es(pfld,pfldtype,pType,pdisabled,pvalue,pbr,ptabindex)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type='"& session("dept") &"EPAY_ES'"
	fsql = fsql & " order by sortfld,cust_code"
	getpay_es = getcodeoption_mul1(conn,fsql,pfldtype,pfld,pType,pdisabled,pvalue,pbr,ptabindex)
end function
'---���y
function getcountry(pType,pcho)
	dim isql
	isql = "select rtrim(coun_code),coun_c from country where markb<>'X' order by coun_code"
	getcountry = getcodeoption(cnn,isql,pType,pcho)
end function
function getcountry1(pType,pcho)
	dim isql
	isql = "select coun_code,coun_c from country where markb<>'X' order by coun_code"
	getcountry1 = getcodeoption(cnn,isql,pType,pcho)
end function
'---sales�笢�Baccount�|�p�Bprocess�ӿ�
function getscode(proles,pType,pcho)
	isql = "select scode,sc_name,scode1 "
	isql = isql & " from vscode_roles"
	isql = isql & " where dept='" &ucase(session("dept"))& "' and syscode = '" &session("syscode")& "'" 
	isql = isql & " and roles='" & proles & "' and branch = '" & session("se_branch") & "'"
	isql = isql & " order by scode1"
	getscode = getcodeoption(cnn,isql,pType,pcho)
end function
'---sales�笢�Baccount�|�p�Bprocess�ӿ�
function getscode2(proles,psyscode,pType,pcho)
	dim isql
	isql = "select distinct scode,sc_name,scode1 "
	isql = isql & " from vscode_roles"
	isql = isql & " where dept='" &ucase(session("dept"))& "' and syscode like '"& psyscode &"'" 
	isql = isql & " and roles in (" & proles & ") and branch = '" & session("se_branch") & "'"
	isql = isql & " order by scode1"
	getscode2 = getcodeoption(cnn,isql,pType,pcho)
end function
'����ץ�D���笢
function getscode1(pType,pcho,ptable)
	dim isql
	isql = "select distinct scode1," & _
		"(case rtrim(scode1) when 'n" & lCase(session("dept")) & "' then '����(�}��Ȥ�)' when 'c" & lCase(session("dept")) & "' then '����(�}��Ȥ�)' when 's" & lCase(session("dept")) & "' then '����(�}��Ȥ�)' when 'k" & lCase(session("dept")) & "' then '����(�}��Ȥ�)'" & _
		" when 'N" & ucase(session("dept")) & "' then '����(�}��Ȥ�)' when 'C" & ucase(session("dept")) & "' then '����(�}��Ȥ�)' when 'S" & ucase(session("dept")) & "' then '����(�}��Ȥ�)' when 'K" & ucase(session("dept")) & "' then '����(�}��Ȥ�)'" & _
		" else isnull(sc_name,'') end) as sc_name" & _
		",(CASE len(substring(scode1, 2, len(scode1))) WHEN 3 THEN '0' + substring(scode1, 2, len(scode1)) ELSE substring(scode1, 2, len(scode1)) END) AS sortscode" & _
		" from "& ptable &" a left join sysctrl.dbo.scode on scode=a.scode1 " & _
		" where scode1 is not null and scode1<>'' order by sortscode"
	getscode1 = getcodeoption(conn,isql,pType,pcho)
end function
'�{��
function getdcscode(pdept,par_type)
    getdcscode = ""
	set rsf = server.CreateObject("Adodb.recordset")
	dim isql
	if par_type="t" then
	    grpid = pdept & "210"
	elseif par_type="e" then
	    grpid = pdept & "220"
	end if
	isql = "select distinct(a.scode),b.sc_name from scode_group a,scode b where a.scode=b.scode"
	isql = isql & " and a.grpclass='" & session("se_branch") & "' and a.grpid='"&grpid&"' and grptype='F'"
	isql = isql & " and (a.end_date is null or a.end_date>='"& date() &" 00:00:00')"
	isql = isql & " order by a.scode"
	'response.write isql & "<BR>"
	rsf.open isql,cnn,1,3
	if not rsf.eof then
	    getdcscode = rsf(0)
	end if
	rsf.close
end function
'�ӿ�
function getprscode(pType,pcho)
	dim isql
	isql = "select distinct(a.scode),b.sc_name from scode_group a,scode b where a.scode=b.scode"
	isql = isql & " and a.grpclass='" & session("se_branch") & "' and (a.grpid='000' or a.grpid='P000' or substring(a.grpid,1,2)='P1' or substring(a.grpid,1,2)='P3' or substring(a.grpid,1,2)='P2' or substring(a.grpid,1,2)='M0')"
	isql = isql & " and (a.end_date is null or a.end_date>='"& date() &" 00:00:00')"
	isql = isql & " order by a.scode"
	getprscode = getcodeoption(cnn,isql,pType,pcho)
end function
'�ӿ�(��ϩ�)
function getprscode1(pType,pcho)
	dim isql
	isql = "select distinct(a.scode),b.sc_name,c.sort,b.sscode"
	isql = isql & " from scode_group a,scode b,branch_code c"
	isql = isql & " where a.scode=b.scode and a.grpclass = c.branch"
	isql = isql & " and ("
	isql = isql & " (a.grpclass IN ('N','C','S','K') and (a.grpid='000' or a.grpid='P000' or substring(a.grpid,1,2)='P1' or substring(a.grpid,1,2)='P2' or substring(a.grpid,1,2)='P3' or substring(a.grpid,1,2)='M0'))"
	isql = isql & " or "
	isql = isql & " (a.grpclass IN ('B','M','T') and substring(a.grpid,1,3)='P3A' and substring(a.grpid,5,1)<>'x')"
	isql = isql & ")"
	isql = isql & " and (a.end_date is null or a.end_date>='"& date() &" 00:00:00')"
	isql = isql & " order by c.sort,b.sscode"
	getprscode1 = getcodeoption(cnn,isql,pType,pcho)
end function
'---picture�s��
function getpic_scode(proles,psyscode,pType,pcho)
	dim isql
	isql = "select distinct scode,sc_name,scode1 "
	isql = isql & " from vscode_roles"
	isql = isql & " where syscode like '"& psyscode &"'" 
	isql = isql & " and roles in (" & proles & ") "
	isql = isql & " order by scode1"
	getpic_scode = getcodeoption(cnn,isql,pType,pcho)
end function
'---��W�ťD��
function getmasterscode(pgrpclass,pscode)
	dim fsql
	dim armasterscode(2)
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select master_scode,upgrpid From scode_group a,grpid b"
	fsql = fsql & " where a.grpclass=b.grpclass and a.grpid=b.grpid"
	fsql = fsql & " and a.grpclass = '"& pgrpclass &"' and a.scode = '"& pscode &"'"
	rst.open fsql,cnn,1,1
	if not rst.eof then
		armasterscode(0) = rst("master_scode")
		armasterscode(1) = rst("upgrpid")
	else
		armasterscode(0) = ""
		armasterscode(1) = ""
	end if
	rst.close
	getmasterscode = armasterscode
end function
'---��խ�
function get_team_scode()
    if (HTProgRight AND 256)<>0 then '�M�׫�
    elseif (HTProgRight AND 128)<>0 then '�@�ťD��
        session("team_man") = ""
        isql = "select distinct a.grpid,a.scode,b.upgrpid From scode_group a inner join grpid b on a.grpclass=b.grpclass and a.grpid=b.grpid"
        isql = isql & " where a.grpclass='"& session("se_branch") &"' and a.end_date>='"& date() &"' and (substring(a.GrpID,1,2)='P1' or a.GrpID='P000' or a.GrpID='000')"
        isql = isql & " and a.GrpID<>'P190' and substring(a.GrpID,5,1)<>'x'"
        isql = isql & " order by a.scode"
        rs1.Open isql,cnn,1,3
        while not rs1.EOF
            session("team_man") = session("team_man") & "'"& rs1("scode") &"',"
            rs1.MoveNext 
        wend
        rs1.Close 
    elseif (HTProgRight AND 64)<>0 then '�G�ťD��
        session("team_man") = ""
        isql = "select distinct a.grpid,a.scode,b.upgrpid From scode_group a inner join grpid b on a.grpclass=b.grpclass and a.grpid=b.grpid"
        isql = isql & " where a.grpclass='"& session("se_branch") &"' and a.end_date>='"& date() &"' and (substring(a.GrpID,1,2)='P1' or a.GrpID='P000')"
        isql = isql & " and a.GrpID<>'P190' and substring(a.GrpID,5,1)<>'x'"
        isql = isql & " order by a.scode"
        rs1.Open isql,cnn,1,3
        while not rs1.EOF
            session("team_man") = session("team_man") & "'"& rs1("scode") &"',"
            rs1.MoveNext 
        wend
        rs1.Close 
    end if
    'response.Write "team_man="& session("team_man") & "<BR>"
end function
'---2016/4/19�W�[�A����`�޳B�{�ǤH���Apsyscode=�t�ΥN�XNTBRT�Apdept=����T�Aproles=����mg_pror=�{��mg_prorm=�D��
function getmgprscode(psyscode,pdept,proles,ptype)
	set rst = server.CreateObject("Adodb.recordset")
	
	mgprscode=""
	
	fsql="select a.scode from scode_roles a inner join scode b on a.scode=b.scode "
	fsql=fsql & " where a.syscode='" & psyscode & "' and a.dept='" & pdept & "' and a.roles='" & proles & "' "
	fsql=fsql & " and (b.end_date is null or b.end_date>='" & date() & "')"
	fsql=fsql & " order by a.sort "
	rst.open fsql,cnn,1,1
	while not rst.eof
	    if ptype="server" then
	        mgprscode=mgprscode & trim(rst("scode")) & "@saint-island.com.tw;"   
	    else
	        mgprscode=mgprscode & trim(rst("scode")) & ";"   
	    end if
	   rst.movenext
	wend
	rst.close
	
	getmgprscode=mgprscode
end function
'---�ϩ�
function getbranch(pwh,pType,pcho)
	dim isql
	isql = "select branch,branchname from branch_code"
	select case pwh
		case "class" isql = isql & " where class='branch'"
		case "mark","showcode" isql = isql & " where "& pwh &"='Y'"
	end select
	isql = isql & " order by sort"
	getbranch = getcodeoption(cnn,isql,pType,pcho)
end function
'---�ϩ�
function getbrancha(pwh,pnobranch,pType,pcho)
	dim isql
	isql = "select branch,branchname from branch_code"
	select case pwh
		case "class" isql = isql & " where class='branch'"
		case "mark","showcode" isql = isql & " where "& pwh &"='Y'"
	end select
	isql = isql & " and branch not in ("& pnobranch &")"
	isql = isql & " order by sort"
	getbrancha = getcodeoption(cnn,isql,pType,pcho)
end function
'---�ϩ�
function getbranchbr(pwh,pType,pcho)
	dim isql
	isql = "select cust_code,code_name from cust_code where code_Type = 'pebr_branch' order by sql"
	'response.write isql & "<BR>"
	getbranchbr = getcodeoption(conn,isql,pType,pcho)
end function
'---�N�z�H
function getagt(pdept,pnobranch,pType,pcho)
	dim isql
	isql = "select agt_no,agt_name from agt"
	isql = isql & " where ("& lcase(pdept) &"end_date is null or "& lcase(pdept) &"end_date='')"
	if pnobranch="Y" then isql = isql & " and (branch is not null and branch<>'')"
	isql = isql & " order by agt_no"
	getagt = getcodeoption(conn,isql,pType,pcho)
end function
'---�M�Q����-�X�M
function getcase1(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type='" & session("Dept") & "Ecase1' order by sortfld"
	getcase1 = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---�ץ����
function getcase_kind(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type = '" & session("Dept") & "ECase_Kind' order by sortfld"
	getcase_kind = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---�ץ���ݳ��
function getcase_source(pType,pcho,pvalue)
	dim fsql
	fsql = "SELECT code,code_name FROM dmp_status where class='case_source' ORDER BY sql"
	getcase_source = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---�y��O
function getslang(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type='pslang' order by sortfld"
	getslang = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---�ꤺ�׮ץ�A��
function getbrp_code(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type='petec_brp' order by sortfld"
	getbrp_code = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---��~�׮ץ�A��
function getexp_code(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type='petec_exp' order by sortfld"
	getexp_code = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---���ޥN�X
function getann_end_code(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code" & _
		" where code_type = '" & session("Dept") & "Eann_end' order by sortfld"
	getann_end_code = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---���ץN�X
function getend_code(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code" & _
		" where code_type = 'ENDCODE' order by sortfld"
	getend_code = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---���׭�]
function getendremark(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code" & _
		" where code_type = 'PEENDREMARK' order by sortfld"
	getendremark = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---���׭�]
function getendremarkp(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code" & _
		" where code_type = 'ENDREMARK' order by sortfld"
	getendremarkp = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'�P�޺���
function get_resp_type(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code" & _
		" where code_type = 'PERESP_TYPE' order by sortfld"
	get_resp_type = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---�o��覡
function getsend_way2(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type = '" & session("Dept") & "Esend_way1' order by sortfld"
	getsend_way2 = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---���覡
function getcase_type(fldname,pdisabled,pType,pcho,pvalue,pbr,ptabindex,ponclick)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type = '" & session("Dept") & "ECase_type1' order by sortfld"
	'getcase_type = showselect5(conn,fsql,pType,pcho,pvalue)
	getcase_type = getcust_code_mul1("radio",fldname,"PECase_type1","","sortfld",pType,pdisabled,pvalue,pbr,ptabindex,ponclick)
end function
'---�q�l�������
function getfile_into_type(pfld,pvalue,pkind)
	getfile_into_type = "<input name='"&pfld&"file_into_type' value='B' type='radio'"
	if pvalue="B" then
		getfile_into_type = getfile_into_type &" checked "
	end if
	getfile_into_type = getfile_into_type & ">�~�O�q��" 
	if pkind="Y" then
		getfile_into_type = getfile_into_type & "<font color='blue'>(" & EcntB & ")</font>"
	end if
	getfile_into_type = getfile_into_type & "<input name='"&pfld&"file_into_type' value='C' type='radio'"
	if pvalue="C" then
		getfile_into_type = getfile_into_type &" checked "
	end if
	getfile_into_type = getfile_into_type & ">�~�O�[���q��"
	if pkind="Y" then
		getfile_into_type = getfile_into_type & "<font color='blue'>(" & EcntC & ")</font>"
	end if
	getfile_into_type = getfile_into_type & "<input name='"&pfld&"file_into_type' value='D' type='radio'"
	if pvalue="D" then
		getfile_into_type = getfile_into_type &" checked "
	end if
	getfile_into_type = getfile_into_type & ">�M�Q�v�����q��"
	if pkind="Y" then
		getfile_into_type = getfile_into_type & "<font color='blue'>(" & EcntD & ")</font>"
	end if
end function
'---������
function getdoc_detail(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type = '" & session("Dept") & "Eatt_DOC'"
	fsql = fsql & " and form_name like '%;"& cgrs &";%' order by sortfld"
	getdoc_detail = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---���M�W�Ǥ���������
function getupdmp_doc1(pType,pcho,pvalue)
	dim fsql
	fsql = "select distinct ref_code,form_name from cust_code"
	fsql = fsql & " where code_type = 'patt_doc'"
	fsql = fsql & " group by ref_code,form_name"
	fsql = fsql & " order by ref_code"
	
	getupdmp_doc1 = showselect7(conn,fsql,pType,pcho,pvalue)
end function
'---���M�W�Ǥ�����
function getupdmp_doc(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type = 'patt_doc'"
	fsql = fsql & " order by sortfld"
	
	getupdmp_doc = showselect7(conn,fsql,pType,pcho,pvalue)
end function
'---�W�Ǥ�����
function getupatt_doc(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type = '" & session("Dept") & "Eatt_doc'"
	if cgrs<>empty then
		fsql = fsql & " and form_name like '%;"& cgrs &";%'"
	end if
	fsql = fsql & " order by sortfld,cust_code"
	
	'response.write "AAA=" & fsql & "<br>"
	'getupatt_doc = showselect5(conn,fsql,pType,pcho,pvalue)
	getupatt_doc = showselect7(conn,fsql,pType,pcho,pvalue)
end function
'---�W�Ǥ�����
function getupatt_doc2(pType,pcho,pvalue)
	dim fsql
	fsql = "select cust_code,code_name from cust_code"
	fsql = fsql & " where code_type = '" & session("Dept") & "Eatt_doc'"
	fsql = fsql & " and mark='C'"	'�Ȥ���
	fsql = fsql & " order by sortfld"
	
	'response.write "AAA=" & fsql & "<br>"
	'getupatt_doc2 = showselect5(conn,fsql,pType,pcho,pvalue)
	getupatt_doc2 = showselect7(conn,fsql,pType,pcho,pvalue)
end function
'����W�Ǥ�����(���W�ǧ@�~���b��)
Function getatt_doc(pType,pcho)
	fsql = "select cust_code,code_name from cust_code where code_type='" & session("dept") & "eatt_doc' order by sortfld"
	getatt_doc = getcodeoption(conn,fsql,pType,pcho)
End Function
'---������ܪ��e��
function getchkform_type(prs_class)
	dim fsql
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select cust_code,mark From cust_code"
	fsql = fsql & " where code_type ='PE95' and cust_code='"& prs_class &"'"
	rst.open fsql,conn,1,1
	if not rst.eof then
		getchkform_type = rst("mark")
	else
		getchkform_type = ""
	end if
	rst.close
end function
'---������ܪ��e��
function getnewold_option(prs_class,pinput_newold) 
	'response.write prs_class & "-" & pinput_newold &"<BR>"
	'response.end
	dim fsql
	dim fwhere
	dim fwhere1
	dim fwhere2
	dim fwhere3
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select cust_code,mark1 From cust_code"
	fsql = fsql & " where code_type ='PE95' and cust_code='"& prs_class &"'"
	'response.write fsql &"<BR>"
	rst.open fsql,conn,1,1
	if not rst.eof then
		'response.write rst("mark1") &"<BR>"
		if pinput_newold = "N" or pinput_newold = "S" then
			if ubound(split(rst("mark1"),";Y;"))>0 then
				fwhere1 = "'N'"
			end if
			if ubound(split(rst("mark1"),";S;"))>0 then
				fwhere2 = "'S'"
			end if
		elseif pinput_newold = "O" then
			if ubound(split(rst("mark1"),";N;"))>0 then
				fwhere3 = "'O'"
			end if
		end if
		if fwhere1<>empty then
			fwhere = fwhere1
		end if
		if fwhere2<>empty and fwhere<>empty then
			fwhere = fwhere & "," & fwhere2
		elseif fwhere2<>empty then
			fwhere = fwhere2
		end if
		if fwhere3<>empty and fwhere<>empty then
			fwhere = fwhere & "," & fwhere3
		elseif fwhere3<>empty then
			fwhere = fwhere3
		end if
		fwhere = " and cust_code in (" & fwhere &")"
		'response.write fwhere &"<BR>"
		'response.end
		getnewold_option = getcust_code(session("dept")&"Enewold",fwhere,"sortfld",false,"N")
	else
		getnewold_option = ""
	end if
	rst.close
end function
'---���׷s�߮�/�D�s�߮׵n���n��ܪ����c����
function getnewold_rs_class(palrs,prs_type,pType,pcho,pmark1,pcolor)
	dim fsql
	fsql = "select cust_code,code_name From cust_code where code_type='"& prs_type &"' and cust_code like 'F%'"
	if pmark1 = "Y" then
		fsql = fsql & " and (mark1 like '%;Y;%' or mark1 like '%;S;%')"
	else
		fsql = fsql & " and mark1 like '%;N;%'"
	end if
	fsql = fsql & " order by cust_code"
	'getaway_flag_rs_class = fsql
	'exit function
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = conn.execute(fsql)
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' selected>�п��</option>"
	end if
	while not tRSa.eof
		if pType then
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
		else
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getnewold_rs_class = innerhtml
	'getnewold_rs_class = getcodeoption(conn,fsql,pType,pcho)
end function
'---������n��ܪ����c����
function getaway_flag_rs_class(palrs,prs_type,pType,pcho,parr)
	dim fsql
	fsql = "select distinct a.cust_code,a.code_name,b.away_flag,a.sortfld From cust_code a,code_exp b"
	fsql = fsql & " where a.code_type ='"& prs_type &"' and a.cust_code like 'F%'"
	fsql = fsql & " and b.rs_type='"& prs_type &"' and b."& palrs &"_flag='Y' and a.cust_code=b.rs_class"
	fsql = fsql & " order by a.sortfld,a.cust_code" 
	'getaway_flag_rs_class = fsql
	'exit function
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = conn.execute(fsql)
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>�п��</option>"
	end if
	while not tRSa.eof
		araway_flag = split(tRSa("away_flag"),",")
		if araway_flag(parr) = "Y" then '�Ĥ@��,�eY��:������ݦ��O�����
			if pType then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if
		elseif araway_flag(parr) = "Y" then '�Ĥ@��,�eY��:������ݦ��O�����
			if pType then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getaway_flag_rs_class = innerhtml
	'getaway_flag_rs_class = getcodeoption(conn,fsql,pType,pcho)
end function
'---���o��N�X-���c����
function getrs_class(palrs,prs_type,pType,pcho)
	dim fsql
	fsql = "select cust_code,code_name from cust_code where code_type='"& prs_type &"'"
	fsql = fsql & " and cust_code in (select rs_class from code_exp where "& lcase(palrs) &"_flag='Y')"
	'fsql = fsql & " and mark1 not like ';X;'"
	if input_rs_class<>empty then
		fsql = fsql & " and cust_code='"& input_rs_class & "'"
	end if
	if submitTask= "A" or (submitTask= "U" and left(stat_code,1)="N") then
		fsql = fsql & " and (end_date is null or end_date>='"& date() &"')"
	end if
	fsql = fsql & " order by sortfld"
	'getrs_class = fsql
	getrs_class = getcodeoption(conn,fsql,pType,pcho)
end function
'---���o��N�X-�שʥN�X
function getrs_code(palrs,prs_type,prs_class,pcountry,psubmittask,pType,pcho)
	fsql = "select rs_code,rs_detail,rs_class from code_exp where " & lcase(palrs) & "_flag ='Y' " & _
		" and rs_type = '" & prs_type & "'" & _
		" and (coun_detail='' or coun_detail is null or coun_detail='ALL' or coun_detail like '%"& pcountry &"%')"
	if psubmittask = "A" or (psubmittask= "U" and left(stat_code,1)="N") then
		fsql = fsql & " and (end_date is null or end_date = '' or end_date > getdate())"
	end if
	if prs_class <> empty then
		fsql = fsql & " and rs_class = '" & prs_class & "'"
	end if
	if (palrs="LR" or palrs="AS") and trim(seq1)<>empty then
		fsql = fsql & " and seq_type = '" & seq1 & "'"
	end if
	fsql = fsql & " order by rs_class,rs_code"
	'getrs_code = fsql
	getrs_code = getcodeoption(conn,fsql,pType,pcho)
end function
'---���o��N�X-�ӿ�ƶ�
function getact_code(palrs,prs_type,prs_class,prs_code,psubmittask,pType,pcho)
	fsql = "select distinct b.act_code, c.code_name ,c.sql"
	fsql = fsql & " from code_exp a, code_actexp b, cust_code c"
	fsql = fsql & " where a." & palrs & "_flag ='Y' "
	fsql = fsql & " and a.rs_type = '" & prs_type & "'"
	fsql = fsql & " and b.cg='" & mid(palrs,1,1) & "' and b.rs = '" & mid(palrs,2,1) & "'"
	fsql = fsql & " and a.sqlno = b.code_sqlno"
	fsql = fsql & " and b.act_code = c.cust_code "
	fsql = fsql & " and c.code_type = 'PEACT_Code'"
	if psubmittask = "A" or (psubmittask= "U" and left(stat_code,1)="N") then
		fsql = fsql & " and (a.end_date is null or a.end_date = '' or a.end_date > getdate())"
		fsql = fsql & " and (b.end_date is null or b.end_date = '' or b.end_date > getdate())"
	end if
	if prs_class <> empty then
		fsql = fsql & " and a.rs_class = '" & prs_class & "'"
	end if
	if prs_code <> empty then
		fsql = fsql & " and a.rs_code = '" & prs_code & "'"
	end if
	fsql = fsql & " order by c.sql"
	'getact_code = fsql
	'response.write fsql
	getact_code = getcodeoption(conn,fsql,pType,pcho)
end function
'---���o��N�X-���שʥN�X
function getrs_codes(palrs,prs_type,prs_class,pcountry,prs_class_flag,psubmittask,pType,pcho)
	fsql = "select rs_code,rs_detail,rs_class from code_exp where " & lcase(palrs) & "_flag ='Y' " & _
		" and rs_type = '" & prs_type & "'" & _
		" and (coun_detail='' or coun_detail is null or coun_detail='ALL' or coun_detail like '%"& pcountry &"%')"
	if psubmittask = "A" then
		fsql = fsql & " and (end_date is null or end_date = '' or end_date > getdate())"
	end if
	if prs_class <> empty then
		fsql = fsql & " and rs_class = '" & prs_class & "'"
	end if
	if (palrs="LR" or palrs="AS") and trim(seq1)<>empty then
		fsql = fsql & " and seq_type = '" & seq1 & "'"
	end if
	fsql = fsql & " and (mark is null or mark='' or mark='B') order by rs_class,rs_code"
	'getrs_code = fsql
	getrs_code = getcodeoption(conn,fsql,pType,pcho)
end function
'---���level
function get_se_grplevel(pscode)
	dim fsql
	set rsf = server.CreateObject("Adodb.recordset")
	fsql = "select b.grpid,b.grplevel from scode_group a "
	fsql = fsql & " inner join grpid b on b.grpclass=a.grpclass and b.grpid=a.grpid "
	fsql = fsql & " and (substring(b.grpid,1,1)='P' or  substring(b.grpid,1,3)='000') "
	fsql = fsql & " where a.scode='"& pscode &"' and a.grpclass='"& session("se_branch") &"'"
	fsql = fsql & " order by grplevel"
	'response.write fsql & "<BR>"
	rsf.open fsql,cnn,1,1
	if not rsf.eof then
		get_se_grplevel = rsf("grplevel")
	else
		get_se_grplevel = "3"
	end if
	rsf.close
end function

'---dn�N�z�H�дڲ��`���D
Function getdn_mark(pType,pcho)
	fsql = "select cust_code,code_name from cust_code where code_type='TDn_mark' order by sortfld"
	getdn_mark = getcodeoption(conn,fsql,pType,pcho)
End Function
'---���O
function getcurrency(pType,pcho)
	fsql = "select currency,currency as nm from ex_rate where class='A'" & _
		" and tr_yy='"& year(date()) &"' and tr_mm='"& month(date()) &"'" & _
		" order by currency"
	getcurrency = getcodeoption(conn,fsql,pType,pcho)
	'getcurrency = fsql
end function
'---�X�f�дڳs���ѹ��O
Function GetCurrencyV(pyear,pmonth,ptype,pcho,Pvalue)
	'response.write "pyear-"&pyear&";"&"pmonth-"&pmonth&";"&"pcho-"&pcho&";"&"pvalue-"&pvalue&";"
	fSQL = "select currency,currency from ex_rate where class='A'"
  	fSQL = fSQL & " and tr_yy='"& pyear &"' and tr_mm='"& pmonth &"' "
	fSQL = fSQL & " order by currency"
  	GetCurrencyV = showselect5(conn,fsql,ptype,pcho,pvalue)
end function

'������
Function getcodeoption(pconn,pSQL,pType,pcho)
'pType:true-->no_name(�N��_�W��), false-->name(�W��)  retrun string
	On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>�п��</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' value1='"& Trim(tRSa(1).value) &"'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
		else
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' value1='"& Trim(tRSa(1).value) &"'>" & Trim(tRSa(1).value) & "</option>"
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getcodeoption=innerhtml
End Function
Function getcodeoption1(pconn,pSQL,pType,pcho,pvalue)
'pType:true-->no_name(�N��_�W��), false-->name(�W��)  retrun string
	On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>�п��</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'"
			if pvalue=Trim(tRSa(0).value) then
				innerhtml=innerhtml & " selected "
			end if
			innerhtml=innerhtml & ">" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
		else
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'"
			if pvalue=Trim(tRSa(0).value) then
				innerhtml=innerhtml & " selected "
			end if
			innerhtml=innerhtml & ">" & Trim(tRSa(1).value) & "</option>"
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getcodeoption1 = innerhtml
End Function
'radio,chkeckbox ������
Function getcodeoption_mul(pconn,pSQL,inputtype,pfldname,pType,pdisabled,pvalue,pbr)
'pType:true-->no_name(�N��_�W��), false-->name(�W��)  retrun string
	'On Error Resume Next
	innerhtml = ""
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	ixcnt = 0
	while not tRSa.eof
		ixcnt = ixcnt + 1
		innerhtml = innerhtml & "<input type='"&inputtype&"' name='"&pfldname&"' value='"& tRSa(0) &"'"& pdisabled
		if pvalue=tRSa(0) then
			innerhtml = innerhtml & " checked "
		end if
		innerhtml = innerhtml & ">"
		if pType=true then
			innerhtml = innerhtml & Trim(tRSa(0).value) & "_"
		end if
		innerhtml = innerhtml & tRSa(1)
		IF pbr="Y" then
			innerhtml = innerhtml & "<br>"
		End IF
		tRSa.MoveNext
	wend
	innerhtml = innerhtml & "<input type='hidden' name='"&pfldname&"cnt' value='"&ixcnt&"'>"
	set tRSa = nothing
	getcodeoption_mul = innerhtml
End Function
'radio,chkeckbox ������
Function getcodeoption_mul1(pconn,pSQL,inputtype,pfldname,pType,pdisabled,pvalue,pbr,ptabindex)
'pType:true-->no_name(�N��_�W��), false-->name(�W��)  retrun string
	'On Error Resume Next
	innerhtml = ""
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	ixcnt = 0
	while not tRSa.eof
		ixcnt = ixcnt + 1
		innerhtml = innerhtml & "<input type='"&inputtype&"' name='"&pfldname&"' value='"& tRSa(0) &"'"& pdisabled
		if pvalue=tRSa(0) then
			innerhtml = innerhtml & " checked "
		end if
		if ptabindex<>empty then
			innerhtml = innerhtml & " tabindex='"& ptabindex &"'"
		end if
		innerhtml = innerhtml & ">"
		if pType=true then
			innerhtml = innerhtml & Trim(tRSa(0).value) & "_"
		end if
		innerhtml = innerhtml & tRSa(1)
		IF pbr="Y" then
			innerhtml = innerhtml & "<br>"
		End IF
		tRSa.MoveNext
	wend
	innerhtml = innerhtml & "<input type='hidden' name='"&pfldname&"cnt' value='"&ixcnt&"'>"
	set tRSa = nothing
	getcodeoption_mul1 = innerhtml
End Function
'radio,chkeckbox ������
Function getcodeoption_mul2(pconn,pSQL,inputtype,pfldname,pType,pdisabled,pvalue,pbr,ptabindex,ponclick)
'pType:true-->no_name(�N��_�W��), false-->name(�W��)  retrun string
	'On Error Resume Next
	innerhtml = ""
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	ixcnt = 0
	while not tRSa.eof
		ixcnt = ixcnt + 1
		innerhtml = innerhtml & "<input type='"&inputtype&"' name='"&pfldname&"' value='"& tRSa(0) &"'"& pdisabled
		if pvalue=tRSa(0) then
			innerhtml = innerhtml & " checked "
		end if
		if ptabindex<>empty then
			innerhtml = innerhtml & " tabindex='"& ptabindex &"'"
		end if
		innerhtml = innerhtml & " onclick='" & ponclick & " "& ixcnt &"'>"
		if pType=true then
			innerhtml = innerhtml & Trim(tRSa(0).value) & "_"
		end if
		innerhtml = innerhtml & tRSa(1)
		IF pbr="Y" then
			innerhtml = innerhtml & "<br>"
		End IF
		tRSa.MoveNext
	wend
	innerhtml = innerhtml & "<input type='hidden' name='"&pfldname&"cnt' value='"&ixcnt&"'>"
	set tRSa = nothing
	getcodeoption_mul2 = innerhtml
End Function

'---------��޸s�M��---begin
'�����޸s�էO
Function getTech_team(pConn,pcho,ptype,pValue)
	'pType:true-->no_name(�էO(�D��)), false-->name(�էO)  retrun string
	fSQL="Select grpid,grpname,(Select sc_name from scode where Master_scode=scode) as master_scode "
	fSQL = fSQL & " from grpid where upgrpid like 'B%' and grpclass='B'"
	'fSQL = fSQL & " and substring(chkcode,5,1)<>'X'"
	fSQL = fSQL & " AND (SUBSTRING(chkcode, 1, 1) = 'Y' OR SUBSTRING(chkcode, 2, 1) = 'Y')"
	On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>�п��</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(fSQL)
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				if Trim(tRSa(2).value)<>empty then
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "(" & Trim(tRSa(2).value) & ")</option>"
				else
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
				end if
			else
				if Trim(tRSa(2).value)<>empty then
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "(" & Trim(tRSa(2).value) & ")</option>"
				else
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
				end if
			end if			
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if			
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getTech_team=innerhtml
End Function
'�����޸s�էO
Function getTech_teamE(pConn,pcho,ptype,pValue)
	'pType:true-->no_name(�էO(�D��)), false-->name(�էO)  retrun string
	fSQL="Select grpid,grpname,(Select sc_name from scode where Master_scode=scode) as master_scode "
	fSQL = fSQL & " from grpid where upgrpid like 'E%' and grpclass='B'"
	fSQL = fSQL & " and substring(chkcode,5,1)<>'X'"
	On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>�п��</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(fSQL)
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				if Trim(tRSa(2).value)<>empty then
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "(" & Trim(tRSa(2).value) & ")</option>"
				else
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
				end if
			else
				if Trim(tRSa(2).value)<>empty then
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "(" & Trim(tRSa(2).value) & ")</option>"
				else
					innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
				end if
			end if			
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if			
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getTech_teamE=innerhtml
End Function

'�����޸s�էO
Function GetTECHscode_team(pConn,pscode)
	fSQL="select grpid from scode_group where scode='"& pscode &"' and grpid like 'B%' "
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(fSQL)
	IF not tRSa.eof then
		GetTECHscode_team=tRSa(0)
	End IF
	set tRSa = nothing
End Function
'-------��޸s�M��---end

'�Ǧ^�n�I���������סApkind:1��Ǧ^��ƪ��סA2��I�������
'pStr:��Ƥ��e�ApLen:��Ƴ̤j���סApCut:�Ǧ^�n�I�������
Function fCutData(pkind,pStr,pLen,pCut)
	if trim(pStr)<>empty then
	else
		exit function
	end if

	fDataLen = 0
	tStr1 = ""
	tStr2 = ""
	pStr = trim(pStr)
	For ixI = 1 To Len(pStr)
		tStr1 = Mid(pStr, ixI, 1)
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then '����r
			tLen = tLen + 2
		Else
			tLen = tLen + 1 '�^�Ʀr
		End If
		
		if pkind="1" then
			if  cint(tLen) <= cint(pLen) then
				tStr2 = tStr2 & tStr1
			end if
		elseif pkind="2" then
			if cint(tLen) > cint(pCut) then
				tStr2 = tStr2 & "..."
				exit for
			end if
			tStr2 = tStr2 & tStr1
		end if
	Next
	fCutData = tStr2
End Function

'---���~������էO
Function getTeam1(ppr_scode)
	set rst = Server.CreateObject("ADODB.Recordset")
	fsql = "select grpid from scode_group where grpclass like 'T%' and grpid like 'T%'" & _
		" and scode='"& trim(ppr_scode) & "'"
	'response.write fsql & "<BR>"
	rst.open fsql,cnn,1,1
	IF not rst.eof then
		getTeam1 = rst("grpid")
		'response.write getteam & "<BR>"
	End IF
	rst.close
End Function
'---���~������էO
Function getTeam2(ppr_scode,pgrpclass,pdept)
	set rst = Server.CreateObject("ADODB.Recordset")
	fsql = "select grpid from scode_group where grpclass like '"&left(pgrpclass,1)&"%' and grpid like '"&left(pdept,1)&"%'" & _
		" and scode='"& trim(ppr_scode) & "'"
	'response.write fsql & "<BR>"
	rst.open fsql,cnn,1,1
	IF not rst.eof then
		getTeam2 = rst("grpid")
		'response.write getteam & "<BR>"
	End IF
	rst.close
End Function

'--�ӿ�H��
function getTeamScode(pvalue,pteam,pleave)

	'���`�H��
	prSQL = "select b.scode,b.sc_name "
	prSQL = prSQL & " from scode b "
	prSQL = prSQL & " where b.scode like 't%' "
	prSQL = prSQL & "   and (b.end_date > getdate() or b.end_date is null) "
	
	prSQL = prSQL & " order by b.sscode "	

	'response.write	"BSQL=" & prSQL
	'response.end
	getTeamScode = showselect5(cnn,prsql,true,"Y",pvalue)

end function

'�P�_���~���O�_���笢
Function getScode_sales(pscode)
	set rst = Server.CreateObject("ADODB.Recordset")

	getScode_sales = false
	tsql = "select a.work_type from sysctrl.dbo.grpid a , sysctrl.dbo.scode_group b "
	tsql = tsql & " where a.grpid = b.grpid "
	tsql = tsql & "   and a.grpclass = b.grpclass "
	tsql = tsql & "   and a.grpclass = '" & session("se_branch") & "'"
	tsql = tsql & "   and b.scode = '" & pscode & "'"
		
	rst.open	tsql,conn,1,1
	if not rst.eof then
		if lcase(rst("work_type")) = "sales" then
			getScode_sales = true
		end if
	end if
	rst.close

end function

function getcgrsname(pcg,prs)
	dim lal
	dim lrs
	if ucase(pcg) = "G" then
		lal = "�x"
	elseif ucase(pcg) = "C" then
		lal = "��"
	elseif ucase(pcg) = "L" or ucase(pcg) = "T" then
		lal = "�p"
	elseif ucase(pcg) = "O" then
		lal = "��"
	elseif ucase(pcg) = "Z" then
		lal = "��"
		if ucase(prs) = "Z" then
			lrs = "��"
		end if
	end if
	if ucase(prs) = "R" then
		lrs = "��"
	elseif ucase(prs) = "S" then
		lrs = "�o"
	end if	
	getcgrsname = lal & lrs
end function

function htmlEncode(pstr)
	pstr = replace(pstr,"&","&amp;")
	pstr = replace(pstr,"<","&lt;")
	pstr = replace(pstr,">","&gt;")
	pstr = replace(pstr,"""","&quot;")
	htmlEncode = pstr
end function
%>
