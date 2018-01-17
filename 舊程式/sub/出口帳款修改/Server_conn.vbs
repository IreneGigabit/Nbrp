<%
function ReplaceData(pdata,pfldnm,pflddata,ptype)
	if ptype="empty" then
		if pflddata<>empty then
			ReplaceData = Replace(pdata, pfldnm, pflddata)
		else
			ReplaceData = Replace(pdata, pfldnm, "")
		end if
	end if
end function

dim fsql
'�ե��ҽs��+�i�� format (barcode)
function formatseqStep(pseq,pseq1,pstep_grade,pjob_sqlno)
	lseqstep = "*SIIPLO-" & session("se_branch") & session("Dept") & "-"
	lseqstep = lseqstep & string(5-len(pseq),"0") & pseq & "-"
	IF pseq1<>"_" then
		lseqstep = lseqstep & pseq1
	End IF
	lseqstep = lseqstep & "-" & string(4-len(pstep_grade),"0") & pstep_grade
	'lseqstep = lseqstep & "-" & string(4-len(pjob_sqlno),"0") & pjob_sqlno & "*"
	lseqstep = lseqstep & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & "*"
	formatseqStep = lseqstep
end function
'�ե��ҽs��+�i�� format (barcode)
function formatseqStep1(pseq,pseq1,pstep_grade,pjob_sqlno)
	lseqstep = "*SIIPLO-" & session("se_branch") & session("Dept") & "-"
	lseqstep = lseqstep & string(5-len(pseq),"0") & pseq & "-"
	IF pseq1<>"_" then
		lseqstep = lseqstep & pseq1
	End IF
	lseqstep = lseqstep & "-" & string(4-len(pstep_grade),"0") & pstep_grade
	lseqstep = lseqstep & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & "*"
	formatseqStep1 = lseqstep
end function
'�ե��ҽs��+�i�� format (barcode)
function formatseqStepPE(pseq_area,pseq,pseq1,pstep_grade,pjob_sqlno)
	lseqstep = "*SIIPLO-" & session("se_branch") & session("Dept") & pseq_area & "-"
	lseqstep = lseqstep & string(5-len(pseq),"0") & pseq & "-"
	IF pseq1<>"_" then
		lseqstep = lseqstep & pseq1
	End IF
	lseqstep = lseqstep & "-" & string(4-len(pstep_grade),"0") & pstep_grade
	'lseqstep = lseqstep & "-" & string(4-len(pjob_sqlno),"0") & pjob_sqlno & "*"
	lseqstep = lseqstep & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & "*"
	formatseqStepPE = lseqstep
end function
'���ˤ����|
function formatscanpath(pseq,pseq1,pstep_grade,pjob_sqlno)
	scanpath = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") & pseq
	'Response.Write scanpath & "<BR>"
	'Response.Write session("scanpathP") & "<BR>"
	scanpath = session("scanpathP") & "/" & left(scanpath,6) & "/"
	'Response.Write scanpath & "<BR>"
	scanfile = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") 
	scanfile = scanfile & pseq & "-"
	if pseq1<>"_" then
		scanfile = scanfile & pseq1 & "-"
	else
		scanfile = scanfile & "-"
	end if
	scanfile = scanfile & string(4-len(pstep_grade),"0") & pstep_grade
	'scanfile = scanfile & "-" & string(4-len(pjob_sqlno),"0") & pjob_sqlno & ".pdf"
	scanfile = scanfile & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & ".pdf"
	formatscanpath = scanpath & scanfile
end function
function formatscanpath1(pseq,pseq1,pstep_grade,pjob_sqlno)
	scanpath = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") & pseq
	'Response.Write scanpath & "<BR>"
	'Response.Write session("scanpathP") & "<BR>"
	scanpath = session("scanpathP2") & "/" & left(scanpath,6) & "/"
	'Response.Write scanpath & "<BR>"
	scanfile = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") 
	scanfile = scanfile & pseq & "-"
	if pseq1<>"_" then
		scanfile = scanfile & pseq1 & "-"
	else
		scanfile = scanfile & "-"
	end if
	scanfile = scanfile & string(4-len(pstep_grade),"0") & pstep_grade
	scanfile = scanfile & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & ".pdf"
	formatscanpath1 = scanpath & scanfile
end function
function formatscanpathPE1(pseq_area,pseq,pseq1,pstep_grade,pjob_sqlno)
	scanpath = session("se_branch") & session("Dept") & pseq_area & "-" & string(5-len(pseq),"0") & pseq
	'Response.Write scanpath & "<BR>"
	'Response.Write session("scanpathPE") & "<BR>"
	'Response.Write session("scanpathPE2") & "<BR>"
	scanpath = session("scanpathPE2") & "/" & left(scanpath,7) & "/"
	'Response.Write scanpath & "<BR>"
	scanfile = session("se_branch") & session("Dept") & pseq_area & "-" & string(5-len(pseq),"0") 
	scanfile = scanfile & pseq & "-"
	if pseq1<>"_" then
		scanfile = scanfile & pseq1 & "-"
	else
		scanfile = scanfile & "-"
	end if
	scanfile = scanfile & string(4-len(pstep_grade),"0") & pstep_grade
	scanfile = scanfile & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & ".pdf"
	formatscanpathPE1 = scanpath & scanfile
end function
function formatscanpath2(pseq,pseq1,pstep_grade,pjob_sqlno,pfile)
	scanpath = session("scan_serverP") & session("scanpathPW") & "\" & left(pfile,6) & "\"
	formatscanpath2 = scanpath & pfile
end function
function formatscanpathf(pseq,pseq1)
	checkpath = session("scan_server") &"\scandoc\"& session("se_branch") & session("Dept")
	scanpath = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") & pseq
	formatscanpathf = checkpath &"\" & left(scanpath,6) 
end function
'���ˤ���ɮצW�� 
function formatscanfilepath(pseq,pseq1,pstep_grade,pjob_sqlno)
	scanfile = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") 
	scanfile = scanfile & pseq & "-"
	if pseq1<>"_" then
		scanfile = scanfile & pseq1 & "-"
	else
		scanfile = scanfile & "-"
	end if
	scanfile = scanfile & string(4-len(pstep_grade),"0") & pstep_grade
	scanfile = scanfile & "-" & string(4-len(pjob_sqlno),"0") & pjob_sqlno & ".pdf"
	formatscanfilepath = scanfile
end function
function formatscanfilepath1(pseq,pseq1,pstep_grade,pjob_sqlno)
	scanfile = session("se_branch") & session("Dept") & "-" & string(5-len(pseq),"0") 
	scanfile = scanfile & pseq & "-"
	if pseq1<>"_" then
		scanfile = scanfile & pseq1 & "-"
	else
		scanfile = scanfile & "-"
	end if
	scanfile = scanfile & string(4-len(pstep_grade),"0") & pstep_grade
	scanfile = scanfile & "-" & string(2-len(pjob_sqlno),"0") & pjob_sqlno & ".pdf"
	formatscanfilepath1 = scanfile
end function
'���oRecordset
function fGetRecordSet(pSQL)
	On Error Resume Next
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa=Conn.execute(pSQL)
	if tRSa.EOF then
		fGetRecordSet = true
	else
		fGetRecordSet = tRSa
	end if
end function
'���o�W��
function getname(pconn,psql)
	getname = ""
	set tRSa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(psql)
	if not tRSa.eof then getname = trim(tRSa(1))
	set tRSa = nothing
end function
'�{��N�X����
function getnow_rs_type(pConn)
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select cust_code from cust_code where code_type='" & session("dept") & "ERS_TYPE'"
	rst.open fsql,pConn,1,1
	if not rst.eof then
		getnow_rs_type = rst(0)
	else
		getnow_rs_type = "PE95"
	end if
	rst.close
end function
'�{��N�X����
function getnow_rs_type_dmp(pConn,pcgrs)
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select mark1 from cust_code where code_type='" & session("dept") & "RS_TYPE' and cust_code='"& pcgrs &"'"
	rst.open fsql,pConn,1,1
	if not rst.eof then
		getnow_rs_type_dmp = rst(0)
	else
		getnow_rs_type_dmp = "P93"
	end if
	rst.close
end function
'�ե��ҽs�� format
function formatseq(branch,dept,pe,seq,seq1,country)
	lseq = branch & ucase(dept)
	if pe="E" then
		lseq = lseq & ucase(pe)
	end if
	lseq = lseq & "-" & seq
	if seq1 <> "_" then
		lseq = lseq & "-" & seq1
	end if
	if country <> empty then
		IF country<>"T" then
			lseq = lseq & "-" & ucase(country)
		End IF
	end if
	formatseq = lseq
end function

'�հ�~�ҥ��ҽs�� format
function formatseqf(pseq,pseq1,pext_flag,pcountry)
Dim lseq
	lseq = session("se_seqtitle") & "-" & pseq
	if pseq1 <> "_" then
		lseq = lseq & "-" & pseq1
	end if
	if pext_flag = "Y" then
		lseq = lseq & "-E"
	end if
	if pcountry <> empty then
		lseq = lseq & " " & ucase(pcountry)
	end if
	formatseqf = lseq
end function

'�ե��ҽs�� format(�Y�O��@��,�h�b�׸��e�[�W"*")
function formatseqf1(pseq,pseq1,pext_flag,pcountry,pcopy_flag)
Dim lseq

	if pcopy_flag = "Y" then
		lseq = "<font color='red'>��</font>"
	end if
	lseq = lseq & session("se_seqtitle") & "-" & pseq
	if pseq1 <> "_" then
		lseq = lseq & "-" & pseq1
	end if
	if pext_flag = "Y" then
		lseq = lseq & "-E"
	end if
	if pcountry <> empty then
		lseq = lseq & " " & ucase(pcountry)
	end if
	formatseqf1 = lseq
end function

'�եN�z�H�s�� format
function formatagent(pagent_no,pagent_no1,pagent_na,pshow_name)
	lagent_no =  pagent_no
	if pagent_no1 <> "_" then
		lagent_no = lagent_no & "-" & pagent_no1
	end if
	if pshow_name = "Y" then
		IF pagent_na<>empty then
			lagent_no = lagent_no & "  " & pagent_na 
		End IF
	end if
	formatagent = lagent_no
end function
'�s�׮ɧ�ץ�s��
function getseq(pkind)
	set rst = server.CreateObject("Adodb.recordset")
	dim fsql
	fsql = "select sql+1 from cust_code where code_type='Z' and cust_code='"& pkind &"'"
	'response.write fsql & "<BR>"
	rst.open fsql,conn,1,1
	getseq = rst(0)
	rst.close
	fsql = " update cust_code set sql=sql+1 where code_type='Z' and cust_code='"& pkind &"'"
	conn.Execute(fsql)
end function
'��ץ�s���ƽX�Afor EU�ڬw
'�ǤJ�̤j���ƽX�A�Y�q1�}�l�s�A����J0
function getseqEU(pkind)
	dim k
	'1~9,A~Y 'Z���i�ϥ� (A:65,Y:89)
	for k=1 to 33
		'response.write asc(pkind) &"<BR>"
		if asc(pkind)=89 then '�Y�OY��w��̫ᤣ�i�A�sZ
			getseqEU = "NO"
			exit function
		elseif asc(pkind)=57 then '�Y�O9(57)�U�@�ӬOA(65)
			pkind = "A"
			if chkseqEU(pkind) = false then
				getseqEU = pkind
				exit function
			end if
		elseif (asc(pkind)>=48 and asc(pkind)<=89) then
			pkind = asc(pkind) + 1
			pkind = chr(pkind)
			if chkseqEU(pkind) = false then
				getseqEU = pkind
				exit function
			end if
		end if
	next
	getseqEU = pkind
	'response.write "getseqEU:"&getseqEU & "<BR>"
end function
function chkseqEU(pseq1)
	chkseqEU = false
	dim fsql
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select * from exp where seq="& seq &" and seq1='"& pseq1 &"'"
	'response.write fsql & "<BR>"
	rst.Open fsql,conn,1,1
	if not rst.EOF then
		chkseqEU = true
	end if
	rst.Close 
end function
'��ץ�s���ƽX�Afor S�s��(���w�s��)
'�ǤJ�̤j���ƽX�A�Y�q1�}�l�s�A����J0
function getseq1_S(pseq,pseq1)
	'1~9,A~Y 'Z,M,C���i�ϥ�
	'response.write asc(pkind) &"<BR>"
	set rst = server.CreateObject("Adodb.recordset")
	dim fsql
	fsql = "select max(seq1) From exp where seq="& pseq &" and seq1='"& pseq1 &"'"
	rst.open fsql,conn,1,1
	if not rst.eof then
		getseq1_S = chr(rst(0)+1)
	end if
	rst.close
end function
'�J�ɮɡA������o��Ǹ�
function getrs_no(pcgrs,premark)
	set rst = server.CreateObject("Adodb.recordset")
	dim fsql
	fsql = "select number from year_num where branch='"& session("se_branch") &"'"
	fsql = fsql & " and dept='"& session("dept") &"E' and num_type='"& pcgrs &"' and num_yy = '" & year(date()) & "'"
	'response.write fsql
	'response.end
	rst.Open fsql,conn,1,1
	if not rst.EOF then
		rs_num = rst("number") + 1
		getrs_no = pcgrs & right(year(date()),2) & string(6-len(rs_num),"0") & rs_num
		
		fsql = "update year_num set number=number+1 where dept='"& session("dept") &"E' and num_type='"& pcgrs &"' and num_yy = '" & year(date()) & "'"
		conn.execute fsql 
	else
		getrs_no = pcgrs & right(year(date()),2) & "000001"
		
		fsql = "insert into year_num(branch,dept,num_type,num_yy,num_mm,number,remark) values"
		fsql = fsql & "('" & session("se_branch") & "','" & session("dept") & "E',"
		fsql = fsql & "'"& pcgrs &"','" & year(date()) & "','_',1,'"& premark &"')"
		'response.write fsql
		'response.end
		conn.execute fsql
	end if
	rst.Close
end function
'�J�ɮɡA������Ǹ�
function getmaxin_no()
	set rst = server.CreateObject("Adodb.recordset")
	dim fsql
	'***2004/1/20�H�e���笢�A�Ǹ��ݥHin_scode,in_no�����ѳ̤j��
	'***2004/1/21�H����笢�A�Ǹ������ܰ�
	fsql = "select max(in_no) as in_no from case_exp"
	if submitTask="A" then 
		fsql = fsql & " where substring(in_no,1,4)='" & year(cdate(date())) & "' and substring(in_no,5,2)='"& string(2-len(month(cdate(date()))),"0") & month(cdate(date())) &"' and substring(in_no,7,2)='"& string(2-len(day(cdate(date()))),"0") & day(cdate(date())) & "'"
	else
		fsql = fsql & " where substring(in_no,1,4)='" & cstr(left(in_no,4)) & "' and substring(in_no,5,2)='"& cstr(string(2-len(mid(in_no,5,2)),"0") & mid(in_no,5,2)) &"' and substring(in_no,7,2)='"& cstr(string(2-len(mid(in_no,7,2)),"0") & mid(in_no,7,2)) & "'"
	end if
	'Response.Write fsql & "<BR>"
	'response.end 
	rst.Open fsql,conn,1,1
	if trim(rst("in_no"))<>empty then
		if submitTask="A" then 
			in_nodate = year(date())&"/"&cstr(string(2-len(month(date())),"0")) & month(date())&"/"&cstr(string(2-len(day(date())),"0")) & day(date())
			if cstr(left(rst("in_no"),4))=cstr(year(in_nodate)) and _
			   cstr(string(2-len(mid(rst("in_no"),5,2)),"0") & mid(rst("in_no"),5,2))=cstr(string(2-len(month(in_nodate)),"0") & month(in_nodate)) and _
			   cstr(string(2-len(mid(rst("in_no"),7,2)),"0") & mid(rst("in_no"),7,2))=cstr(string(2-len(day(in_nodate)),"0") & day(in_nodate)) then
				in_no = rst("in_no")+1
			else
				in_no = cstr(year(date())) & cstr(string(2-len(month(date())),"0") & month(date())) & cstr(string(2-len(day(date())),"0") & day(date())) & "001"
			end if
		else
			in_no = request("in_no")
		end if
	else
		in_no = year(date()) & string(2-len(month(date())),"0") & month(date()) & string(2-len(day(date())),"0") & day(date()) & "001"
	end if
	rst.Close
end function
function getmaxcase_no()
	dim fsql
	set rst = server.CreateObject("Adodb.recordset")
	nowmonth = cstr(string(2-len(month(date())),"0")) & month(date())
	fsql = "select max(case_no) as case_no from case_exp"
	fsql = fsql & " where substring(case_no,1,3)='E"& year(date())-2000 &"'"
	fsql = fsql & " and substring(case_no,4,2)='"& nowmonth &"'"
	if session("scode")="admin" then
		Response.Write "test--"& fsql & "<Br>"
		'response.end
	end if
	rst.open fsql,conn,1,1
	case_no = trim(rst("case_no"))
	'case_no = "E1112999"  'test  E1131042
	if case_no<>empty then
		if date()>="2011/1/1" then '2011���褸�~���X
			'nowmonth = "12"  'test
			if cstr(mid(case_no,2,2))="99" and cstr(mid(case_no,4,2))="12" then
				ycase_no = "11"
				mcase_no = "01"
			end if
			if session("scode")="admin" then
				'Response.Write ycase_no & "=" & cstr(year(date())-2000) & "<Br>"
				'Response.Write mcase_no & "=" & cstr(nowmonth) & "<Br>"
				'Response.Write rst("case_no") & "<Br>"
				
			end if
			'if (ycase_no=cstr(year(date())-2000)) and (mcase_no=cstr(nowmonth)) then
			if (cstr(mid(case_no,2,2))=cstr(year(date())-2000)) and (cstr(mid(case_no,4,2))=cstr(nowmonth)) _
			or (cstr(mid(case_no,2,2))=cstr(year(date())-2000)) and (cstr(mid(case_no,4,1))=cstr(cdbl(nowmonth))) _
			or (cstr(mid(case_no,2,2))=cstr(year(date())-2000)) and (cstr(mid(case_no,4,1))="A") _
			or (cstr(mid(case_no,2,2))=cstr(year(date())-2000)) and (cstr(mid(case_no,4,1))="B") _
			or (cstr(mid(case_no,2,2))=cstr(year(date())-2000)) and (cstr(mid(case_no,4,1))="C") then
				if mid(case_no,6,8)="999" or (cdbl(asc(mid(case_no,4,1)))>=65 and cdbl(asc(mid(case_no,4,1)))<=90) then
					'�C��u��999���i�ϥΡA�W�L999������h�אּ
					'�@���E�� E1103999-->E1131000-->E1131001
					'�Q���12�� E1110999-->E11A1000-->E11A1001�AE1111999-->E11B1000-->E11B1001�AE1112999-->E11C1000-->E11C1001
					if (cstr(mid(case_no,2,2))=cstr(year(date())-2000)) and (cstr(mid(case_no,4,1))=cstr(int(nowmonth))) then 
						getmaxcase_no = mid(case_no,1,4) & cstr(int(mid(case_no,5))+1)
					else
						if int(nowmonth)=10 then
							getmaxcase_no = cstr(mid(case_no,2,2)) & "A1000"
						elseif int(nowmonth)=11 then
							getmaxcase_no = cstr(mid(case_no,2,2)) & "B1000"
						elseif int(nowmonth)=12 then
							getmaxcase_no = cstr(mid(case_no,2,2)) & "C1000"
						else
							getmaxcase_no = cstr(mid(case_no,2,2)) & cstr(mid(case_no,5,1)) & "1000"
						end if
					end if
				else
					if cdbl(asc(cstr(mid(case_no,4,1))))>=65 and cdbl(asc(cstr(mid(case_no,4,1))))<=90 then  'A~Z
						'�W�L999����
						getmaxcase_no = cstr(mid(case_no,2,3)) & cstr(int(mid(case_no,4))+1)
					else
						getmaxcase_no = cstr(int(mid(case_no,2))+1)
					end if
				end if
			else
				getmaxcase_no = cstr(year(date())-2000) & cstr(string(2-len(month(date())),"0")) & cstr(month(date())) & "001"
			end if
			if session("scode")="admin" then
				Response.Write getmaxcase_no & "<br>"
				'Response.End 
			end if
		else
			nowmonth = cstr(string(2-len(month(date())),"0")) & month(date())
			'Response.Write mid(case_no,2,2) & "/"&mid(case_no,4,2) & "<Br>"
			if (cstr(mid(case_no,2,2))=cstr(year(date())-1911)) and (cstr(mid(case_no,4,2))=cstr(nowmonth)) then
				getmaxcase_no = cstr(int(mid(case_no,2))+1)
			else
				getmaxcase_no = cstr(year(date())-1911) & cstr(string(2-len(month(date())),"0")) & cstr(month(date())) & "001"
			end if
		end if
	else
		if date()>="2011/1/1" then '2011���褸�~���X
			getmaxcase_no = cstr(year(date())-2000) & cstr(string(2-len(month(date())),"0")) & cstr(month(date())) & "001"
		else
			getmaxcase_no = cstr(year(date())-1911) & cstr(string(2-len(month(date())),"0")) & cstr(month(date())) & "001"
		end if
	end if
	rst.close
end function
'���´�s��
function getteam(pscode,pgrpid)
	dim fsql
	set rst = server.CreateObject("Adodb.recordset")
	fsql = "select grpid from scode_group where grpclass='"& session("se_branch") &"'"
	fsql = fsql & " and scode='"& pscode &"' and (grpid like '"& pgrpid &"%' or grpid='000') order by grpid"
	'response.write fSQL&"<br>"
	rst.open fsql,cnn,1,1
	if not rst.eof then
		getteam = rst("grpid")
	else
		getteam = ""
	end if
	rst.close
end function
'�ˬd���ʮשʮ�,�O�_�|�@�֧�ʮץ󪬺A
'�Y�ǤJ���שʪ������ץ󪬺A����,�h������s�D��
'�Y�ǤJ���שʪ������ץ󪬺A�L��,
'      �h�P�_�Ӷi�׭ק�e���שʬO�_���ץ󪬺A, 
'      �Y�L,�h���έק�ץ�D��
'      �Y��,�h�ݧ��̫�@�Ӧ��ץ󪬺A�� 
function update_case_stat(pconn,pseq,pseq1,pstep_grade,prs_type,prs_class,prs_code,pact_code,pcase_stat,prs_sqlno)
	Dim tisql
	Dim rst
	Dim tcase_stat,trs_type,trs_class,trs_code,tnow_grade
	Dim fcase_stat,frs_type,frs_class,frs_code,fstep_grade
	Dim ocase_stat,ors_type,ors_class,ors_code,ostep_grade
	
	'response.write "step_grade="& pstep_grade &"<BR>"
	'response.write "case_stat="& pcase_stat &"<BR>"
	fcase_stat = ""
	frs_type =""
	frs_class=""
	frs_code =""
	fstep_grade=""
	set rst = server.CreateObject("Adodb.recordset")
	if pcase_stat <> "" then
		'��ܶǤJ�שʦ������ץ󪬺A
		'Response.Write "aaa" & "<br>"
		fstep_grade=trim(pstep_grade)
		frs_type = trim(prs_type)
		frs_class = trim(prs_class)
		frs_code = trim(prs_code)
		fact_code = trim(pact_code)
		fcase_stat = trim(pcase_stat)
	else
		'��ܶǤJ�שʨS�������ץ󪬺A
		'�ˬdnow_grade�O�_����pstep_grade,�p�G�O�h���̫�@�����ץ󪬺A
		tisql = "select now_grade from exp "
		tisql = tisql & " where seq = " & pseq
		tisql = tisql & "   and seq1 = '" & pseq1 & "'"
		'Response.write tisql & "<Br>"
		'Response.end
		rst.Open tisql,pconn,1,1
		if not rst.EOF then
			tnow_grade = trim(rst("now_grade"))
		end if
		rst.close
		
		if tnow_grade = pstep_grade then
			tisql = "select * from step_exp a"
			tisql = tisql & " where a.seq = " & pseq
			tisql = tisql & "   and a.seq1 = '" & pseq1 & "'"
			tisql = tisql & "   and a.step_grade <> '" & pstep_grade & "'"
			tisql = tisql & "   and (a.case_stat is not null and a.case_stat<>'') "
			tisql = tisql & " order by step_grade desc"
			'Response.write tisql & "<Br>"
			rst.Open tisql,pconn,1,1
			'Response.Write "���o�e�@���v�T�ץ󪬺A���i�קǸ� <br>" & tisql & "<br>" & "err--" & err.number & "<br>"
			if not rst.eof then
				frs_type = trim(rst("rs_type"))
				frs_class = trim(rst("rs_class"))
				frs_code = trim(rst("rs_code"))
				fact_code = trim(rst("act_code"))
				fcase_stat = trim(rst("case_stat"))
				fstep_grade = trim(rst("step_grade"))
			end if
			rst.close
		else
			'�Ynow_grade <> pstep_grade, �h���έק�
			frs_type = ""
			frs_class = ""
			frs_code = ""
			fact_code = ""
			fcase_stat = ""
			fstep_grade = ""
		end if
	end if
	
	'response.write "step_grade="& fstep_grade &"<BR>"
	if fstep_grade <> empty then
		tisql = " select now_stat,now_arcase_type,now_arcase_class,now_arcase,now_act_code,now_grade "
		tisql = tisql & " from exp "
		tisql = tisql & " where seq = " & pseq
		tisql = tisql & "   and seq1 = '" & pseq1 & "'"
		rst.Open tisql,pconn,1,1
		if not rst.EOF then
			ocase_stat = trim(rst("now_stat"))
			ors_type = trim(rst("now_arcase_type"))
			ors_class = trim(rst("now_arcase_class"))
			ors_code = trim(rst("now_arcase"))
			oact_code = trim(rst("now_act_code"))
			ostep_grade = trim(rst("now_grade"))
		end if
		rst.close
		tisql = "update exp set  "
		tisql = tisql & " now_stat= '" & fcase_stat & "'"
		tisql = tisql & " ,now_arcase_type='" & frs_type & "'"
		tisql = tisql & " ,now_arcase_class='" & frs_class & "'"
		tisql = tisql & " ,now_arcase='" & frs_code & "'"
		tisql = tisql & " ,now_act_code='" & fact_code & "'"
		tisql = tisql & " ,now_grade='" & fstep_grade & "'"
		tisql = tisql & " where seq = " & pseq
		tisql = tisql & "   and seq1 = '" & pseq1 & "'"
		
		'Response.Write "�ק�ץ󪬺A:" & tisql & "<br>"
		'Response.end
		pconn.execute tisql 
		if err.number<>0 then
			update_case_stat = true
		end if
		
		'�g�J��slog
		if ocase_stat <> fcase_stat then
			call insert_exp_rec_log(pconn,pseq,pseq1,"exp","now_stat",ocase_stat,fcase_stat,prgid,prs_sqlno)
		end if
		if ors_type <> frs_type then
			call insert_exp_rec_log(pconn,pseq,pseq1,"exp","now_arcase_type",ors_type,frs_type,prgid,prs_sqlno)
		end if
		if ors_class <> frs_class then
			call insert_exp_rec_log(pconn,pseq,pseq1,"exp","now_arcase_class",ors_class,frs_class,prgid,prs_sqlno)
		end if
		if ors_code <> frs_code then
			call insert_exp_rec_log(pconn,pseq,pseq1,"exp","now_arcase",ors_code,frs_code,prgid,prs_sqlno)
		end if
		if oact_code <> fact_code then
			call insert_exp_rec_log(pconn,pseq,pseq1,"exp","now_act_code",oact_code,fact_code,prgid,prs_sqlno)
		end if
		if ostep_grade <> fstep_grade then
			call insert_exp_rec_log(pconn,pseq,pseq1,"exp","now_grade",ostep_grade,fstep_grade,prgid,prs_sqlno)
		end if
	end if
	set rst = nothing
end function	
'---�J ext_rec_log
function insert_exp_rec_log(pconn,pseq,pseq1,ptable_name,pfield_name,povalue,pnvalue,pprgid,prs_sqlno)
	tsql = "insert into exp_rec_log (branch,seq,seq1,table_name,field_name,ovalue,nvalue,"
	tsql = tsql & "tran_date,tran_scode,prgid,rs_sqlno) values("
	tsql = tsql & "'"& session("se_branch") &"','"& pseq & "','" & pseq1 & "','"& ptable_name &"',"
	tsql = tsql & "'" & pfield_name & "','" & povalue & "','" & pnvalue & "',"
	tsql = tsql & "getdate(),'" & session("se_scode") & "','" & pprgid & "','" & prs_sqlno & "'"
	tsql = tsql & ")"
	'Response.Write "insert ext_rec_log <br>" & tsql & "<br>"	
	'response.end
	conn.execute tsql	
end function

'�J�ɮɡA����� null
function chkdatenull(pvalue)
	if trim(pvalue)<>empty then
		chkdatenull = "'"& trim(pvalue) &"'"
	else
		chkdatenull = "null"
	end if
End Function
'�J�ɮɡA����� null
function chkdatenull2(pvalue)
	if trim(pvalue)<>empty then
		chkdatenull2 = "'"& FormatDateTime(pvalue,2) &" "& string(2-len(hour(pvalue)),"0") & hour(pvalue) &":"& string(2-len(minute(pvalue)),"0") & minute(pvalue) &":"& string(2-len(second(pvalue)),"0") & second(pvalue) &"'"
	else
		chkdatenull2 = "null"
	end if
End Function
'�J�ɮɡA����� null for informix
function chkdatenull3(pvalue)
	if trim(pvalue)<>empty then
		chkdatenull3 = "'"& string(2-len(month(pvalue)),"0") & month(pvalue) &"/"& string(2-len(day(pvalue)),"0") & day(pvalue) &"/"& year(pvalue) & "'"
		'chkdatenull3 = "'"& FormatDateTime(pvalue,2) &" "& string(2-len(hour(pvalue)),"0") & hour(pvalue) &":"& string(2-len(minute(pvalue)),"0") & minute(pvalue) &":"& string(2-len(second(pvalue)),"0") & second(pvalue) &"'"
	else
		chkdatenull3 = "null"
	end if
End Function
'�J�ɮɡAchar���ť�
function chkcharnull(pvalue)
	if trim(pvalue)<>empty then
		pvalue = replace(pvalue,"'","��")
		pvalue = replace(pvalue,"""","��")
		pvalue = replace(pvalue,"&","��")
		chkcharnull = "'"& trim(pvalue) &"'"
	else
		chkcharnull = "''"
	end if
End Function
'�J�ɮɡAchar���ť�
function chkcharnull2(pvalue)
	if trim(pvalue)<>empty then
		chkcharnull2 = "'"& trim(pvalue) &"'"
	else
		chkcharnull2 = "''"
	end if
End Function
'�J�ɮɡAchar��null for informix
function chkcharnull3(pvalue)
	if trim(pvalue)<>empty then
		chkcharnull3 = "'"& trim(pvalue) &"'"
	else
		chkcharnull3 = "null"
	end if
End Function
'�J�ɮɡAchar���ť� �S��r
function chkcharnull4(pvalue)
	if trim(pvalue)<>empty then
		ar1 = split(pvalue,"&#")
		if ubound(ar1)>0 then
			pvalue = pvalue
		else
			pvalue = Replace(pvalue,"&","��")
		end if
		pvalue = Replace(pvalue,"'","`")		
		pvalue = replace(pvalue,"'","��")
		pvalue = replace(pvalue,"""","��")
		chkcharnull4 = "'"& trim(pvalue) &"'"
	else
		chkcharnull4 = "''"
	end if
End Function
'�J�ɮɡAnumber��0
function chknumzero(pvalue)
	if trim(pvalue)<>empty then
		chknumzero = pvalue
	else
		chknumzero = 0
	end if
End Function

function chkempty(p1)
	if trim(p1)<>empty then
		chkempty = "'" & p1 & "'"
	else
		chkempty = "''"
	end if
end function

function formatcgrs(pcg,prs)
	Dim tcg,trs
	tcg = ""
	if ucase(pcg) = "A" then
		tcg="�N"
	elseif ucase(pcg) = "L" then
		tcg ="�p"
	elseif ucase(pcg) = "O" then
		tcg = "��"
	elseif ucase(pcg) = "Z" then
		tcg = "��"
	elseif ucase(pcg) = "T" then
		tcg ="�p"
	elseif ucase(pcg) = "C" then
		tcg ="��"	
	elseif ucase(pcg) = "G" then
		tcg ="�x"	
	end if
	
	if ucase(prs) = "R" then
		trs="��"
	elseif ucase(prs) = "S" then
		trs ="�o"
	elseif ucase(prs) = "Z" then
		trs ="��"
	end if

	formatcgrs = tcg & trs
	
end function


'�Ǧ^�T�w���צr��(����vbscript:left)
'pStr:��Ƥ��e
'pLen:��Ƴ̤j����,�Y�ǤJ0�h�Ǧ^��ƪ���
Function fMid(pStr,pLen)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	fDataLen = 0
	tStr1 = ""
	tStr2 = ""
	if pStr<>empty then
		For ixI = 1 To Len(pStr)
			tStr1 = Mid(pStr, ixI, 1)
			tCod = Asc(tStr1)
			If tCod >= 128 Or tCod < 0 Then
				tLen = tLen + 2
			Else
				tLen = tLen + 1
			End If
			
			if  tLen <= pLen then
				tStr2 = tStr2 & tStr1
			end if
		Next
		if  tLen > pLen then
			tStr2 = tStr2 & "..."
		end if
		fMid = tStr2
	else
		fMid = ""
	end if
End Function

'�Ǧ^�T�w���צr��,�q�Y�@�r��}�l(����vbscript:mid)
'pStr:��Ƥ��e
'pLen:��Ƴ̤j����,�Y�ǤJ0�h�Ǧ^��ƪ���
Function fLeft(pStr,pLen)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	fDataLen = 0
	tStr1 = ""
	tStr2 = ""
	tStr3 = ""
	rStr2=fMid(pStr,pLen)
	i=len(rStr2)+1
	For ixI = i To Len(pStr)
		tStr1 = Mid(pStr, ixI, 1)
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then
			tLen = tLen + 2
		Else
			tLen = tLen + 1
		End If
		
		if  tLen <= pLen then
			tStr3 = tStr3 & tStr1
		end if
	Next

	fLeft = tStr3
End Function
'�Ǧ^�T�w���צr��(����vbscript:right)
'pStr:��Ƥ��e
'pLen:��Ƴ̤j����,�Y�ǤJ0�h�Ǧ^��ƪ���
Function fRight(pStr,pLen)
	Dim ixI 
	Dim tStr1
	Dim tCod
	Dim tLen
	Dim i
	fDataLen = 0
	tStr1 = ""
	tStr2 = ""
	i=Len(pStr)
	For ixI = i To len(pLen) step -1
		
		tStr1 = Mid(pStr, ixI, 1)
		
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then
			tLen = tLen + 2
		Else
			tLen = tLen + 1
		End If
		
		if  tLen <= pLen then
			tStr2 = tStr1 & tStr2
		end if
	Next
	fRight = tStr2
End Function
'�ˬd�ި����,�O�_�ݭn�{���鴣�e�ܶg��
function scheck_week_ctrldate(pctrl_type,pdate)
	Dim gCtrl_type,tCtrl_Type , tday , tdate,tbaseday
	
	scheck_week_ctrldate = pdate
	
	if pdate="" or isnull(pdate) then
		exit function
	end if
	
	'�ثe�u���ۺ޴���,�ӿ����,�ӿ�o������ݭn�{���鴣�e�ܶg��
	gCtrl_Type = "B,D,F"
	
	tctrl_type = left(pctrl_type,1)

	if instr(gCtrl_type,tCtrl_Type)=0 then
		exit function
	end if

	tbaseday = 5 '�g��	
	tday = weekday(pdate,vbMonday)
	
	'�{�g��.�P��~�ݹB��
	if (tday = 6) or (tday = 7) then
		tdate = DateAdd("d",-(tday-tbaseday),pdate)
		scheck_week_ctrldate = tdate
	end if
end function

'�ˬd�ި����,�O�_�ݭn�{���鴣�e�ܶg��
function scheck_week_ctrldate(pctrl_type,pdate)
	Dim gCtrl_type,tCtrl_Type , tday , tdate,tbaseday
	
	scheck_week_ctrldate = pdate
	
	if pdate="" or isnull(pdate) then
		exit function
	end if
	
	'�ثe�u���ۺ޴���,�ӿ����,�ӿ�o������ݭn�{���鴣�e�ܶg��
	gCtrl_Type = "B,D,F"
	
	tctrl_type = left(pctrl_type,1)

	if instr(gCtrl_type,tCtrl_Type)=0 then
		exit function
	end if

	tbaseday = 5 '�g��	
	tday = weekday(pdate,vbMonday)
	
	'�{�g��.�P��~�ݹB��
	if (tday = 6) or (tday = 7) then
		tdate = DateAdd("d",-(tday-tbaseday),pdate)
		scheck_week_ctrldate = tdate
	end if
end function


'**********************��bserver_savelog.vbs
'�s�W Log �ɡA�A�Ω� log table ���� ud_flag�Bud_date�Bud_scode�Bprgid �o������
'ptable�Gex:step_imp �n�s�W�� step_imp_log �h�ǤJ  step_imp
'pkey_filed�Gkey �����W�١A�p���h�����ХΡF�j�}
'pkey_value�G�P pkey_field �ۤ��t�X�A�p���h�����ХΡF�j�}
function insert_log_tablexxx(pconn,pud_flag,pprgid,ptable,pkey_field,pkey_value)
	dim tisql
	dim tfield_str
	dim ar_key_field
	dim ar_key_value
	dim wsql
	dim ti
	
	set tRS = Server.CreateObject("ADODB.Recordset")
	
	tfield_str = ""
	
	tisql = "select b.name from sysobjects a, syscolumns b "
	tisql = tisql & " where a.id = b.id  and a.name = '" & ptable & "' and a.xtype='U' "
	tisql = tisql & " order by b.colid "
	
	tRS.open tisql,pconn,1,1
	while not tRS.eof
		tfield_str = tfield_str & tRS("name") & ","
		tRS.MoveNext
	wend
		
	tRS.close
	
	tfield_str = left(tfield_str,len(tfield_str) - 1)
	
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
	
	tisql = "insert into " & ptable & "_log(ud_flag,ud_date,ud_scode,prgid," & tfield_str & ")"
	tisql = tisql & "select '" & pud_flag & "',getdate(),'" &session("scode") & "',"
	tisql = tisql & "'" &pprgid& "'," & tfield_str
	tisql = tisql & " from " & ptable
	tisql = tisql & " where 1 = 1 "
	tisql = tisql & wsql
	pconn.execute tisql
	
	set tRS = nothing
end function
'---�M�Q����
function getcode_dnp_class(pdn_case_type,pType,pcho)
	fsql = "select cust_code,code_name from cust_codef where code_type = '" & pdn_case_type & "' order by sortfld"
	getcode_dnp_class = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---�дڤ��e�y�z���o
function getcode_dnp(pdn_case_type,pType,pcho)
'	if session("scode")="admin" then
'	fsql="select top 100 dn_case,dn_case_detail from code_dnp "
'	else
	fsql="select dn_case,dn_case_detail from code_dnp "
'	end if
	fsql = fsql & " where 1=1 "
	IF pdn_case_type<>empty then
		fsql = fsql & " and dn_case_type='"& pdn_case_type &"'"
	End IF
	IF pdn_case_class<>empty then	
		 fsql = fsql & " and dn_case_class='"& pdn_case_class &"'"
	End IF
	
	getcode_dnp = getcodeoption(conn,fsql,pType,pcho)
end function

function check_in_date(pseq,pseq1,pflddate,pfldvalue)
	check_in_date = true
	'response.write "in_date="& pfldvalue &"<BR>"
	if pfldvalue<>empty then
		check_in_date = false
	else
		check_in_date = true
		pseq = formatseq(session("se_branch"),session("Dept"),"",pseq,pseq1,"")
		select case pflddate
			case "in_date"
				pfldname = "�߮פ��"
			case "apply_date"
				pfldname = "�ӽФ��"
			case "apply_no"
				pfldname = "�ӽи�"
		end select
		'Response.Write pseq & pfldname & "���L��ƽгq����T�� !!!"
		%>
		<html><head><meta http-equiv='Content-Type' content='text/html; charset=big5'><title></title></head>
		<body>
		<script language='vbscript'>
			'on error resume next			msgbox "<%=pseq%>" & "<%=pfldname%>" &"�L��ƽгq����T�� !!!"
		</script>
		</body>
		</html>
		<%
	end if
	'response.write "check_in_date="& check_in_date &"<BR>"
end function

'�g�Jexch_rec_log
function insert_exch_rec_log(pconn,pexch_no,pfield_name,povalue,pnvalue,pprgid)
	tsql = "insert into exch_rec_log (exch_no,field_name,ovalue,nvalue,tran_date,tran_scode,prgid)values("
	tsql = tsql & pexch_no & ",'" & pfield_name & "','" & povalue & "','" & pnvalue & "',"
	tsql = tsql & "getdate(),'" & session("se_scode") & "','" & pprgid & "')"
	'Response.Write "insert imt_rec_log <br>" & tsql & "<br>"	
	pconn.execute tsql
end function
%>
