<!--#INCLUDE FILE="../sub/Server_cbx.vbs" -->
<%
'取得Recordset
function GetRecordSet(pConn,pSQL)
	On Error Resume Next
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa=pConn.execute(pSQL)
	if tRSa.EOF then
		fGetRecordSet = true
	else
		set fGetRecordSet = tRSa
	end if
end function
'組本所編號 format
function formatseq(seq,seq1,ext_flag,country,dept)
	lseq = dept & "-" & seq
	if seq1 <> "_" then
		lseq = lseq & "-" & seq1
	end if
	if ext_flag = "Y" then
		lseq = lseq & "-E"
	end if
	if country <> empty then
		lseq = lseq & " " & ucase(country)
	end if
	formatseq = lseq
end function
'組成案件名稱
function formatAppl(pcappl_name,peappl_name)
	if pcappl_name <> empty then
		if len(pcappl_name) > 20 then
			formatAppl = mid(pcappl_name,1,20) &  "..."
		else
			formatAppl = pcappl_name
		end if
	else
		if  peappl_name <> empty and len(peappl_name) > 20 then
			formatAppl = mid(peappl_name,1,20) &  "..."
		else
			formatAppl = peappl_name
		end if
	end if
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
	fMid = tStr2
End Function
'傳回要截取的欄位長度，pkind:1表傳回資料長度，2表截取的資料
'pStr:資料內容，pLen:資料最大長度，pCut:傳回要截取的資料
Function fCutData(pkind,pStr,pLen,pCut)
	if trim(pStr)<>empty then
	else
		exit function
	end if

	fDataLen = 0
	tStr1 = ""
	tStr2 = ""
	
	For ixI = 1 To Len(pStr)
		tStr1 = Mid(pStr, ixI, 1)
		tCod = Asc(tStr1)
		If tCod >= 128 Or tCod < 0 Then '中文字
			tLen = tLen + 2
		Else
			tLen = tLen + 1 '英數字
		End If
		
		if pkind="1" then
			if  tLen <= pLen then
				tStr2 = tStr2 & tStr1
			end if
		elseif pkind="2" then
			if tLen>pCut then
				tStr2 = tStr2 & "..."
				exit for
			end if
			tStr2 = tStr2 & Mid(pStr, ixI, 1)
		end if
	Next
	fCutData = tStr2
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
'---國籍
function getcountry(pType,pcho)
	isql = "select coun_code,coun_c from country where coun_code<>'T' and markf<>'I' and markf<>'X' order by coun_code"
	getcountry = getcodeoption(cnn,isql,pType,pcho)
end function
'---代理人種類
function getagt()
	isql = "select cust_code,code_name from cust_code where code_type = 'Agent_Type' order by sortfld"
	getagt = getcodeoption(conn,isql,false,"Y")
end function
'---代理人種類 radio box
function getagtr(pfield,pprgid,pagent_type,psub)
	dim i
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql="select cust_code,code_name from cust_code where code_type = 'Agent_Type'"
'	if lcase(trim(pprgid))<>"agent18" and lcase(trim(pprgid))<>"agent19" _
'	and lcase(trim(pprgid))<>"agent41" then 
'		isql = isql & " and (mark<>'X' or mark is null or mark='')"
'	end if
	isql = isql & " order by sortfld"
	getagtr = ""
	i = 0
	RSget.open isql,conn,1,1
	while not RSget.eof
		getagtr = getagtr & "<input type='radio' name='"& pfield &"' value='" & RSget("cust_code") & "' "
		if trim(ucase(RSget("cust_code"))) = trim(ucase(pagent_type)) then
			getagtr = getagtr & " checked "
		end if
		if psub<>empty then
			getagtr = getagtr & " onclick='"& psub &" """& RSget("cust_code") &""""&"'"
		end if
		getagtr = getagtr & Qdisabled & " >" & RSget("code_name")
		i = i + 1
		RSget.movenext
	wend
	RSget.close
	getagtr = getagtr & "<input type='hidden' name='"& pfield &"_cnt' value='"& i &"'>"
end function
'---代理人種類 check box 'agent61.asp
function getagt1(pfield,pprgid,pagent_type,psub,pbr)
	dim i
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql="select cust_code,code_name from cust_code where code_type = 'Agent_Type'"
	if lcase(trim(pprgid))<>"agent18" and lcase(trim(pprgid))<>"agent19" _
	and lcase(trim(pprgid))<>"agent41" and lcase(trim(pprgid))<>"agent61" then 
		isql = isql & " and (mark<>'X' or mark is null or mark='')"
	end if
	isql = isql & " order by sortfld"
	getagt1 = ""
	i = 0
	RSget.open isql,conn,1,1
	while not RSget.eof
		getagt1 = getagt1 & "<input type=checkbox name='"& pfield &"' value='"& RSget(0) &"'"
		if trim(ucase(RSget("cust_code"))) = trim(ucase(pagent_type)) then
			getagt1 = getagt1 & " checked "
		end if
		if psub<>empty then
			getagt1 = getagt1 & " onclick='"& psub &" """& RSget("cust_code") &""""&"'"
		end if
		getagt1 = getagt1 & pdisabled &">"& RSget(1)
		if pbr="Y" and (i mod 2)=0 then getagt1 = getagt1 & "<br>"
		i = i + 1
		RSget.movenext
	wend
	RSget.close
	getagt1 = getagt1 & "<input type='hidden' name='"& pfield &"_cnt' value='"& i &"'>"
end function
'---信用等級
function getcredit()
	isql = "select cust_code,code_name from cust_code where code_type = 'credit' order by cust_code"
	getcredit = getcodeoption(conn,isql,false,"Y")
end function
'---列印備註
function getpmark()
	isql = "select cust_code,code_name from cust_code where code_type = 'Pmark' order by cust_code"
	getpmark = getcodeoption(conn,isql,false,"Y")
end function
'---折扣方式
function getdis_type()
	isql = "select cust_code,code_name from cust_code where code_type = 'dis_type' order by cust_code"
	getdis_type = getcodeoption(conn,isql,false,"Y")
end function
'---部門
function getDept()
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept' and (mark is null or mark='') order by cust_code"
	getDept = getcodeoption(conn,isql,true,"Y")
end function
'---部門1
function getDept1n()
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept1' and (mark is null or mark='') order by sortfld"
	getDept1n = getcodeoption(conn,isql,true,"Y")
end function
'---部門1
function getDept2n()
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept1' and substring(mark1,1,1)='Y' order by sortfld"
	getDept2n = getcodeoption(conn,isql,true,"Y")
end function
'---部門1
function getDept2()
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept2' and substring(mark1,1,1)='Y' order by sortfld"
	getDept2 = getcodeoption(conn,isql,true,"Y")
end function
'---部門1--多選
function getDept1(pname,pbr,pdisabled,pwhere)
	'pname:欄位名稱， pbr:Y:表換行
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept1' and (mark is null or mark='')" & _
		pwhere & " order by sortfld"
	RSget.open isql,conn,1,1
	i = 1
	while not RSget.eof
		getDept1 = getDept1 & "<input type=checkbox name='"& pname&RSget(0) &"' value='"& RSget(0) &"' "& pdisabled &">"& RSget(1)
		if pbr="Y" and (i mod 2)=0 then getDept1 = getDept1 & "<br>"
		i = i + 1
		RSget.movenext
	wend
	RSget.close
	'getDept1 = getcodeoption(conn,isql,true,"Y")
end function
'---部門1--多選
function getDept1_1(pname,pbr,pdisabled,pwhere)
	'pname:欄位名稱， pbr:Y:表換行
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept1' and (mark is null or mark='')" & _
		pwhere & " order by sortfld"
	RSget.open isql,conn,1,1
	i = 0
	while not RSget.eof
		i = i + 1
		getDept1_1 = getDept1_1 & "<input type=checkbox name='"& pname&i &"' value='"& RSget(0) &"' "& pdisabled &">"& RSget(1)
		if pbr="Y" and (i mod 2)=0 then getDept1 = getDept1 & "<br>"
		RSget.movenext
	wend
	RSget.close
	getDept1_1 = getDept1_1 & "<input type='hidden' name='"& pname &"cnt' value='"& i &"'>"
	'getDept1_1 = getcodeoption(conn,isql,true,"Y")
end function
'---部門--名稱
function getDept1_nm(pvalue)
	'pname:欄位名稱， pbr:Y:表換行
	if trim(pvalue)=empty then
		getDept1_nm = ""
		exit function
	end if
	dim i
	set RSget = Server.CreateObject("ADODB.Recordset")
	pwhere = ""
	arpvalue = split(pvalue,";")
	for i=0 to ubound(arpvalue)-1
		pwhere = pwhere & "'"& arpvalue(i) &"',"
	next
	pwhere = " and cust_code in (" & left(pwhere,len(pwhere)-1) & ")"
	isql = "select cust_code,code_name from cust_code where code_type = 'Dept1' "
	isql = isql & pwhere & " order by sortfld"
	RSget.open isql,conn,1,1
	getDept1_nm = ""
	i = 0
	while not RSget.eof
		i = i + 1
		getDept1_nm = getDept1_nm & RSget("code_name") & "、"
		RSget.movenext
	wend
	RSget.close
	if i>0 then
		getDept1_nm = left(getDept1_nm,len(getDept1_nm)-1)
	end if
	'getDept1_nm = isql
end function
'---報導方式--多選
function getnews_type(pname,pbr,pdisabled,pwhere)
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql = "select cust_code,code_name from cust_code where code_type = 'Anews_type'" & _
		pwhere & " order by sortfld"
	RSget.open isql,conn,1,1
	i = 1
	while not RSget.eof
		getnews_type = getnews_type & "<input type='checkbox' name='"& pname &"' value='"& RSget(0) &"' "& pdisabled &">"& RSget(1)
		if pbr="Y" and (i mod 2)=0 then getnews_type = getnews_type & "<br>"
		i = i + 1
		RSget.movenext
	wend
	RSget.close
	getnews_type = getnews_type & "<input type='hidden' name='"& pname &"_cnt' value='"& i-1 &"'>"
end function
'---報導方式--radio
function getnews_type1(pfield,pprgid,pagent_type,psub)
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql = "select cust_code,code_name from cust_code where code_type = 'Anews_type'" & _
		" order by sortfld"
	getnews_type1 = ""
	RSget.open isql,conn,1,1
	while not RSget.eof
		getnews_type1 = getnews_type1 & "<input type='radio' name='"& pfield &"' value='" & RSget("cust_code") & "' "
		if trim(ucase(RSget("cust_code"))) = trim(ucase(pagent_type)) then
			getnews_type1 = getnews_type1 & " checked "
		end if
		if psub<>empty then
			getnews_type1 = getnews_type1 & " onclick='"& psub &" """& RSget("cust_code") &""""&"'"
		end if
		getnews_type1 = getnews_type1 & Qdisabled & " >" & RSget("code_name")
		RSget.movenext
	wend
	RSget.close
end function
'---聯絡方式
function getatt_type()
	isql = "select cust_code,code_name from cust_code where code_type = 'ATT_TYPE' order by sortfld"
	getatt_type = getcodeoption1(conn,isql,true,"Y","N")
end function
'---聯絡方式--多選
function getatt_type1(pname,pbr,pdisabled)
	'pname:欄位名稱， pbr:Y:表換行
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql = "select cust_code,code_name from cust_code where code_type = 'ATT_TYPE' order by sortfld"
	RSget.open isql,conn,1,1
	i = 1
	while not RSget.eof
		getatt_type1 = getatt_type1 & "<input type=checkbox name='"& pname &"' value='"& RSget(0) &"' "& pdisabled &">"& RSget(1)
		if pbr="Y" and (i mod 2)=0 then getatt_type1 = getatt_type1 & "<br>"
		i = i + 1
		RSget.movenext
	wend
	RSget.close
end function
'---代理人結束代碼
function getend_code()
	isql = "select cust_code,code_name from cust_code where code_type = 'AEND_CODE' order by sortfld"
	getend_code = getcodeoption(conn,isql,true,"Y")
end function
'---聯絡人結束代碼
function getatend_code()
	isql = "select cust_code,code_name from cust_code where code_type = 'ATEND_CODE' order by sortfld"
	getatend_code = getcodeoption(conn,isql,true,"Y")
end function
'---聯絡人
function getAgent_att(value1,value2)
	isql = "select seqno,att_name from agt_att where dept like '%"&ucase(Session("dept"))&"%'"
	isql = isql&" and agent_no='"&value1&"'"
	isql = isql&" and agent_no1='"&value2&"'"
	getAgent_att = getcodeoption(conn,isql,false,"Y")
end function
'---分案別
function gettbr_type()
	isql = "select cust_code,code_name from cust_code where code_type='TBR_TYPE'"
	gettbr_type = getcodeoption(conn,isql,true,"Y")
end function
'---進專分案別
function getpbr_type()
	isql = "select cust_code,code_name from cust_code where code_type='PBR_TYPE' order by sortfld"
	getpbr_type = getcodeoption(conn,isql,true,"Y")
end function
'---洲別
function getcoun_area()
	isql = "select cust_code,code_name from cust_code where code_type='TCoun_Area' order by sortfld"
	getcoun_area = getcodeoption(conn,isql,true,"Y")
end function
'---持卷人員組別
function getbranch_ag(pgrpid,pType,pcho)
	set RSget = Server.CreateObject("ADODB.Recordset")
	isql = "select distinct a.grpclass,(select grpname from grpid where grpclass=a.grpclass and grpid='000') as grpclassnm,a.grpid,a.grpname" & _
		" from grpid a where (a.grpclass in ('A','B','L','T','TT','TP','TS') or a.grpid='B000' or a.grpid='D000' or a.grpid='D100' or a.grpid='D200')" & _
		" and substring(a.grpid,1,3)<>'P3A'" 
	if trim(pgrpid)<>empty then
		isql = isql & " and a.grpid='"& pgrpid &"'"
	end if
	isql = isql & " order by a.grpclass,a.grpid"
	htmlstr = ""
	if pcho = "Y" then htmlstr = htmlstr & "<option value='' style='color:blue' selected>請選擇</option>"
	RSget.open isql,cnn,1,1
	while not RSget.eof
		if RSget("grpid")="000" then
			htmlstr = htmlstr & "<option value='"& RSget("grpclass") &"_"& RSget("grpid") &"'>"& RSget("grpclassnm") & "</option>"
		else
			htmlstr = htmlstr & "<option value='"& RSget("grpclass") &"_"& RSget("grpid") &"'>"& RSget("grpclassnm") & RSget("grpname") & "</option>"
		end if
		RSget.movenext
	wend
	RSget.close
	getbranch_ag = htmlstr
	'getbranch_ag = getcodeoption(cnn,isql,pType,pcho)
end function
'---持卷人員
function getscode_ag(pgrpid,pType,pcho)
	isql = "select distinct a.scode,b.sc_name,b.sscode" & _
		" from scode_group a,scode b " & _
		" where a.scode=b.scode and (a.grpclass in ('A','B','L','T','TT','TP','TS') or a.grpid='B000' or a.grpid='D000' or a.grpid='D100' or a.grpid='D200')" & _
		" and substring(a.grpid,1,3)<>'P3A'" 
	if trim(pgrpid)<>empty then
		isql = isql & " and a.grpid='"& pgrpid &"'"
	end if
	isql = isql & " order by b.sscode"
	'getscode_ag= isql
	getscode_ag = getcodeoption(cnn,isql,pType,pcho)
end function
'---代理人收文代碼大類
function getagt_r_code(ptype,pcho)
	isql = "select cust_code,code_name from cust_code where code_type='AGT_R_CODE'" & _
		" and substring(cust_code,1,1)='_' order by sortfld"
	getagt_r_code = getcodeoption(conn,isql,ptype,pcho)
end function
'---代理人收發文種類
function getagtrs_kind(pfield,ponclick,psubmitTask)
	isql = "select cust_code,code_name from cust_code where code_type='AGTRS_KIND'" & _
		" order by sortfld"
	set RSget = Server.CreateObject("ADODB.Recordset")
	i = 0
	RSget.open isql,conn,1,1
	while not RSget.eof
		getagtrs_kind = getagtrs_kind & "<input type=radio name='"& pfield &"' value='"& RSget("cust_code") &"'"
		if ponclick<>empty then getagtrs_kind = getagtrs_kind & " onclick='"& ponclick &"'"
		if psubmitTask<>"A" then getagtrs_kind = getagtrs_kind & " disabled "
		getagtrs_kind = getagtrs_kind & ">"& RSget("code_name") &"&nbsp;&nbsp;"
		i = i + 1
		RSget.movenext
	wend
	RSget.close
	getagtrs_kind = getagtrs_kind & "<input type=hidden name=rs_kindcnt value='"& i &"'>"
	getagtrs_kind = getagtrs_kind & "<input type=hidden name=hrs_kind>"
	'getagtrs_kind = getcodeoption(conn,isql,true,"Y")
end function
'---代理人發文代碼大類
function getagt_s_code(ptype,pcho)
	isql = "select cust_code,code_name from cust_code where code_type='AGT_S_CODE'" & _
		" and substring(cust_code,1,1)='_' order by sortfld"
	getagt_s_code = getcodeoption(conn,isql,ptype,pcho)
end function
'---發文正本單位
function getsend_cl(ptype,pcho)
	isql = "select cust_code,code_name from cust_code where code_type='ASEND_CL'" & _
		" order by sortfld"
	getsend_cl = getcodeoption(conn,isql,ptype,pcho)
end function
'---會談類別
function getrec_type()
	isql = "select cust_code,code_name from cust_code where code_type='REC_TYPE'"
	getrec_type = getcodeoption(conn,isql,false,"Y")
end function
'---會談類別 radio
function getrec_typer(pvalue)
	getrec_typer = ""
	set RecRS = Server.CreateObject("ADODB.Recordset")
	isql = "select cust_code,code_name from cust_code where code_type='REC_TYPE'"
	set RecRS = conn.execute(isql)
	while not RecRS.eof
		getrec_typer = getrec_typer & "<input type=radio name='rec_type' value='"& RecRS(0) &"'"
		if pvalue=RecRS(0) then getrec_typer = getrec_typer & " checked "
		getrec_typer = getrec_typer & ">"& RecRS(1)
		RecRS.movenext
	wend
	set RecRS = nothing
end function

'取得來文方式
function getReceive_Way(pconn,pagrs)
	itemstr = ""
	set RecRS = Server.CreateObject("ADODB.Recordset")
	isql = "select * from cust_code where code_type = 'TREC_WAY' " 
	isql = isql & " and mark1 like '%" & pagrs & "%'"
	isql = isql & " order by sortfld"
	
	set RecRS = pConn.execute(isql)
	while not RecRS.eof
		itemstr = itemstr & RecRS("Code_Name") & ";"
		RecRS.movenext
	wend
	
	set RecRS = nothing
	
	if trim(itemstr) <> "" then
		itemstr = mid(itemstr,1,len(itemstr) - 1)
	end if
	getReceive_Way = itemstr
end function
function getReceive_WayID(pconn,pagrs)
	idstr = ""
	set RecRS = Server.CreateObject("ADODB.Recordset")
	isql = "select * from cust_code where code_type = 'TREC_WAY' " 
	isql = isql & " and mark1 like '%" & pagrs & "%'"
	isql = isql & " order by sortfld"
	
	set RecRS = pConn.execute(isql)
	while not RecRS.eof
		idstr = idstr & RecRS("Cust_code") & ";"
		RecRS.movenext
	wend
	
	set RecRS = nothing
	
	if trim(idstr) <> "" then
		idstr = mid(idstr,1,len(idstr) - 1)
	end if
	getReceive_WayID = idstr
end function

'取得代理人發文方式
function getagt_send_Way(pconn,pagrs)
	itemstr = ""
	set RecRS = Server.CreateObject("ADODB.Recordset")
	isql = "select * from cust_code where code_type = 'ASEND_WAY'" 
	isql = isql & " order by sortfld"
	
	set RecRS = pConn.execute(isql)
	while not RecRS.eof
		itemstr = itemstr & RecRS("Code_Name") & ";"
		RecRS.movenext
	wend
	
	set RecRS = nothing
	
	if trim(itemstr) <> "" then
		itemstr = mid(itemstr,1,len(itemstr) - 1)
	end if
	getagt_send_Way = itemstr
end function
function getagt_send_WayID(pconn,pagrs)
	idstr = ""
	set RecRS = Server.CreateObject("ADODB.Recordset")
	isql = "select * from cust_code where code_type = 'ASEND_WAY'" 
	isql = isql & " order by sortfld"
	
	set RecRS = pConn.execute(isql)
	while not RecRS.eof
		idstr = idstr & RecRS("Cust_code") & ";"
		RecRS.movenext
	wend
	
	set RecRS = nothing
	
	if trim(idstr) <> "" then
		idstr = mid(idstr,1,len(idstr) - 1)
	end if
	getagt_send_WayID = idstr
end function

'抓取資料
Function getcodeoption(pconn,pSQL,pType,pcho)
'pType:true-->no_name(代號_名稱), false-->name(名稱)  retrun string
	On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>請選擇</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
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
	getcodeoption=innerhtml
End Function
'抓取資料
Function getcodeoption1(pconn,pSQL,pType,pcho,phavename)
	'pType:true-->no_name(代號_名稱), false-->name(名稱)  retrun string
	On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>請選擇</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if phavename="Y" then
			if pType then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if
		else
			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "</option>"
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing
	getcodeoption1=innerhtml
End Function
%>
