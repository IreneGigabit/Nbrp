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
'組本所編號+進度 format (barcode)
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
'組本所編號+進度 format (barcode)
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
'組本所編號+進度 format (barcode)
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
'掃瞄文件路徑
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
'掃瞄文件檔案名稱 
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
'取得Recordset
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
'取得名稱
function getname(pconn,psql)
	getname = ""
	set tRSa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(psql)
	if not tRSa.eof then getname = trim(tRSa(1))
	set tRSa = nothing
end function
'現行代碼種類
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
'現行代碼種類
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
'組本所編號 format
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

'組國外所本所編號 format
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

'組本所編號 format(若是改作案,則在案號前加上"*")
function formatseqf1(pseq,pseq1,pext_flag,pcountry,pcopy_flag)
Dim lseq

	if pcopy_flag = "Y" then
		lseq = "<font color='red'>＊</font>"
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

'組代理人編號 format
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
'新案時抓案件編號
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
'抓案件編號副碼，for EU歐洲
'傳入最大的副碼，若從1開始編，請轉入0
function getseqEU(pkind)
	dim k
	'1~9,A~Y 'Z不可使用 (A:65,Y:89)
	for k=1 to 33
		'response.write asc(pkind) &"<BR>"
		if asc(pkind)=89 then '若是Y表已到最後不可再編Z
			getseqEU = "NO"
			exit function
		elseif asc(pkind)=57 then '若是9(57)下一個是A(65)
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
'抓案件編號副碼，for S新案(指定編號)
'傳入最大的副碼，若從1開始編，請轉入0
function getseq1_S(pseq,pseq1)
	'1~9,A~Y 'Z,M,C不可使用
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
'入檔時，抓取收發文序號
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
'入檔時，抓取交辦序號
function getmaxin_no()
	set rst = server.CreateObject("Adodb.recordset")
	dim fsql
	'***2004/1/20以前改營洽，序號需以in_scode,in_no抓取當天最大值
	'***2004/1/21以後改營洽，序號不需變動
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
		if date()>="2011/1/1" then '2011改抓西元年後兩碼
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
					'每月只有999號可使用，超過999給號原則改為
					'一月到九月 E1103999-->E1131000-->E1131001
					'十月到12月 E1110999-->E11A1000-->E11A1001，E1111999-->E11B1000-->E11B1001，E1112999-->E11C1000-->E11C1001
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
						'超過999給號
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
		if date()>="2011/1/1" then '2011改抓西元年後兩碼
			getmaxcase_no = cstr(year(date())-2000) & cstr(string(2-len(month(date())),"0")) & cstr(month(date())) & "001"
		else
			getmaxcase_no = cstr(year(date())-1911) & cstr(string(2-len(month(date())),"0")) & cstr(month(date())) & "001"
		end if
	end if
	rst.close
end function
'抓組織群組
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
'檢查當更動案性時,是否會一併更動案件狀態
'若傳入的案性的對應案件狀態有值,則直接更新主檔
'若傳入的案性的對應案件狀態無值,
'      則判斷該進度修改前的案性是否有案件狀態, 
'      若無,則不用修改案件主檔
'      若有,則需找到最後一個有案件狀態的 
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
		'表示傳入案性有對應案件狀態
		'Response.Write "aaa" & "<br>"
		fstep_grade=trim(pstep_grade)
		frs_type = trim(prs_type)
		frs_class = trim(prs_class)
		frs_code = trim(prs_code)
		fact_code = trim(pact_code)
		fcase_stat = trim(pcase_stat)
	else
		'表示傳入案性沒有對應案件狀態
		'檢查now_grade是否等於pstep_grade,如果是則找到最後一筆有案件狀態
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
			'Response.Write "取得前一筆影響案件狀態之進度序號 <br>" & tisql & "<br>" & "err--" & err.number & "<br>"
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
			'若now_grade <> pstep_grade, 則不用修改
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
		
		'Response.Write "修改案件狀態:" & tisql & "<br>"
		'Response.end
		pconn.execute tisql 
		if err.number<>0 then
			update_case_stat = true
		end if
		
		'寫入更新log
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
'---入 ext_rec_log
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

'入檔時，日期給 null
function chkdatenull(pvalue)
	if trim(pvalue)<>empty then
		chkdatenull = "'"& trim(pvalue) &"'"
	else
		chkdatenull = "null"
	end if
End Function
'入檔時，日期給 null
function chkdatenull2(pvalue)
	if trim(pvalue)<>empty then
		chkdatenull2 = "'"& FormatDateTime(pvalue,2) &" "& string(2-len(hour(pvalue)),"0") & hour(pvalue) &":"& string(2-len(minute(pvalue)),"0") & minute(pvalue) &":"& string(2-len(second(pvalue)),"0") & second(pvalue) &"'"
	else
		chkdatenull2 = "null"
	end if
End Function
'入檔時，日期給 null for informix
function chkdatenull3(pvalue)
	if trim(pvalue)<>empty then
		chkdatenull3 = "'"& string(2-len(month(pvalue)),"0") & month(pvalue) &"/"& string(2-len(day(pvalue)),"0") & day(pvalue) &"/"& year(pvalue) & "'"
		'chkdatenull3 = "'"& FormatDateTime(pvalue,2) &" "& string(2-len(hour(pvalue)),"0") & hour(pvalue) &":"& string(2-len(minute(pvalue)),"0") & minute(pvalue) &":"& string(2-len(second(pvalue)),"0") & second(pvalue) &"'"
	else
		chkdatenull3 = "null"
	end if
End Function
'入檔時，char給空白
function chkcharnull(pvalue)
	if trim(pvalue)<>empty then
		pvalue = replace(pvalue,"'","’")
		pvalue = replace(pvalue,"""","”")
		pvalue = replace(pvalue,"&","＆")
		chkcharnull = "'"& trim(pvalue) &"'"
	else
		chkcharnull = "''"
	end if
End Function
'入檔時，char給空白
function chkcharnull2(pvalue)
	if trim(pvalue)<>empty then
		chkcharnull2 = "'"& trim(pvalue) &"'"
	else
		chkcharnull2 = "''"
	end if
End Function
'入檔時，char給null for informix
function chkcharnull3(pvalue)
	if trim(pvalue)<>empty then
		chkcharnull3 = "'"& trim(pvalue) &"'"
	else
		chkcharnull3 = "null"
	end if
End Function
'入檔時，char給空白 特殊字
function chkcharnull4(pvalue)
	if trim(pvalue)<>empty then
		ar1 = split(pvalue,"&#")
		if ubound(ar1)>0 then
			pvalue = pvalue
		else
			pvalue = Replace(pvalue,"&","＆")
		end if
		pvalue = Replace(pvalue,"'","`")		
		pvalue = replace(pvalue,"'","’")
		pvalue = replace(pvalue,"""","”")
		chkcharnull4 = "'"& trim(pvalue) &"'"
	else
		chkcharnull4 = "''"
	end if
End Function
'入檔時，number給0
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
		tcg="代"
	elseif ucase(pcg) = "L" then
		tcg ="聯"
	elseif ucase(pcg) = "O" then
		tcg = "其"
	elseif ucase(pcg) = "Z" then
		tcg = "本"
	elseif ucase(pcg) = "T" then
		tcg ="聯"
	elseif ucase(pcg) = "C" then
		tcg ="客"	
	elseif ucase(pcg) = "G" then
		tcg ="官"	
	end if
	
	if ucase(prs) = "R" then
		trs="收"
	elseif ucase(prs) = "S" then
		trs ="發"
	elseif ucase(prs) = "Z" then
		trs ="收"
	end if

	formatcgrs = tcg & trs
	
end function


'傳回固定長度字串(類似vbscript:left)
'pStr:資料內容
'pLen:資料最大長度,若傳入0則傳回資料長度
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

'傳回固定長度字串,從某一字串開始(類似vbscript:mid)
'pStr:資料內容
'pLen:資料最大長度,若傳入0則傳回資料長度
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
'傳回固定長度字串(類似vbscript:right)
'pStr:資料內容
'pLen:資料最大長度,若傳入0則傳回資料長度
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
'檢查管制期限,是否需要逢假日提前至週五
function scheck_week_ctrldate(pctrl_type,pdate)
	Dim gCtrl_type,tCtrl_Type , tday , tdate,tbaseday
	
	scheck_week_ctrldate = pdate
	
	if pdate="" or isnull(pdate) then
		exit function
	end if
	
	'目前只有自管期限,承辦期限,承辦發文期限需要逢假日提前至週五
	gCtrl_Type = "B,D,F"
	
	tctrl_type = left(pctrl_type,1)

	if instr(gCtrl_type,tCtrl_Type)=0 then
		exit function
	end if

	tbaseday = 5 '週五	
	tday = weekday(pdate,vbMonday)
	
	'逢週六.周日才需運算
	if (tday = 6) or (tday = 7) then
		tdate = DateAdd("d",-(tday-tbaseday),pdate)
		scheck_week_ctrldate = tdate
	end if
end function

'檢查管制期限,是否需要逢假日提前至週五
function scheck_week_ctrldate(pctrl_type,pdate)
	Dim gCtrl_type,tCtrl_Type , tday , tdate,tbaseday
	
	scheck_week_ctrldate = pdate
	
	if pdate="" or isnull(pdate) then
		exit function
	end if
	
	'目前只有自管期限,承辦期限,承辦發文期限需要逢假日提前至週五
	gCtrl_Type = "B,D,F"
	
	tctrl_type = left(pctrl_type,1)

	if instr(gCtrl_type,tCtrl_Type)=0 then
		exit function
	end if

	tbaseday = 5 '週五	
	tday = weekday(pdate,vbMonday)
	
	'逢週六.周日才需運算
	if (tday = 6) or (tday = 7) then
		tdate = DateAdd("d",-(tday-tbaseday),pdate)
		scheck_week_ctrldate = tdate
	end if
end function


'**********************改在server_savelog.vbs
'新增 Log 檔，適用於 log table 中有 ud_flag、ud_date、ud_scode、prgid 這些欄位者
'ptable：ex:step_imp 要新增至 step_imp_log 則傳入  step_imp
'pkey_filed：key 值欄位名稱，如有多個欄位請用；隔開
'pkey_value：與 pkey_field 相互配合，如有多個欄位請用；隔開
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
'---專利種類
function getcode_dnp_class(pdn_case_type,pType,pcho)
	fsql = "select cust_code,code_name from cust_codef where code_type = '" & pdn_case_type & "' order by sortfld"
	getcode_dnp_class = showselect5(conn,fsql,pType,pcho,pvalue)
end function
'---請款內容描述取得
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
				pfldname = "立案日期"
			case "apply_date"
				pfldname = "申請日期"
			case "apply_no"
				pfldname = "申請號"
		end select
		'Response.Write pseq & pfldname & "期無資料請通知資訊部 !!!"
		%>
		<html><head><meta http-equiv='Content-Type' content='text/html; charset=big5'><title></title></head>
		<body>
		<script language='vbscript'>
			'on error resume next			msgbox "<%=pseq%>" & "<%=pfldname%>" &"無資料請通知資訊部 !!!"
		</script>
		</body>
		</html>
		<%
	end if
	'response.write "check_in_date="& check_in_date &"<BR>"
end function

'寫入exch_rec_log
function insert_exch_rec_log(pconn,pexch_no,pfield_name,povalue,pnvalue,pprgid)
	tsql = "insert into exch_rec_log (exch_no,field_name,ovalue,nvalue,tran_date,tran_scode,prgid)values("
	tsql = tsql & pexch_no & ",'" & pfield_name & "','" & povalue & "','" & pnvalue & "',"
	tsql = tsql & "getdate(),'" & session("se_scode") & "','" & pprgid & "')"
	'Response.Write "insert imt_rec_log <br>" & tsql & "<br>"	
	pconn.execute tsql
end function
%>
