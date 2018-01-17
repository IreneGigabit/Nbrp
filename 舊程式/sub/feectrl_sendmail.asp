
<%
Begin_time=now


Dim StrToList,StrToList1, Sender,subject,body,BCCStrToList

function feectrl_onclick()
   mail_flag="N"
   
   '檢查cust_code.code_type=Z and cust_code.cust_code=PEfc_mail_date的日期
   fsql="select form_name from cust_code where code_type='Z' and cust_code='PEfc_mail_date' and (end_date is null or end_date>=getdate()) "
   rsi.open fsql,conn,1,1
   if not rsi.eof then
      mail_date=trim(rsi("form_name"))
      if cdate(mail_date)<>date() then 
         mail_flag="Y"
      end if   
   end if
   rsi.close
   
   if mail_flag="Y" then	
        '期限超過1~2天
		task=sendmail("2")
		if task="Y" then
		    DoSendMail subject,body
		end if  
		'期限超過3天
		task1=sendmail("3")
		if task1="Y" then
		    DoSendMail subject,body
		end if   
		'修改cust_code的Email通知日期
		if task="Y" or task1="Y" then
		    usql="update cust_code set form_name='" & date() & "' where code_type='Z' and cust_code='PEfc_mail_date' "
		    conn.execute usql
		end if 
   end if		
end function
'ptodo="2"期限超過1-2天，通知到區所主管,ptodo="3"期限超過3天，通知到執委
function Sendmail(ptodo)
	pdate=dateadd("d",-3,date())
	sql="select c.branch,a.scode1,count(*) as cnt "
	sql = sql & " ,(select sc_name from sysctrl.dbo.scode where scode=a.scode1) as sc_name "
	sql = sql & " ,(select branchname from sysctrl.dbo.branch_code where branch=c.branch) as branchname "
    sql = sql & " from todo_exp t inner join exp as a on a.seq=t.seq and a.seq1=t.seq1 " 
    sql = sql & " inner join step_exp c on c.seq=t.seq and c.seq1=t.seq1 and c.rs_sqlno = t.rs_sqlno"
	sql = sql & " where t.syscode='" & session("syscode") & "'"	
	sql = sql & " and t.dowhat in ('SC_TRF') and t.job_status in ('NN','NX')"
	if ptodo="2" then
	   sql = sql & " and dateadd(month,1,c.step_date)>'" & pdate & "' and dateadd(month,1,c.step_date)<'" & date() & "'"
	elseif ptodo="3" then
	   sql = sql & " and dateadd(month,1,c.step_date)<='" & pdate & "'"
	end if   
	sql = sql & " group by c.branch,a.scode1 "
	sql = sql & " order by c.branch,convert(int,substring(a.scode1,2,5)) "
	'Response.Write sql&"<br>"
	'Response.End
	rsi.open sql,conn,1,1
	StrToList1=""	'副本
	send_mail="N"
	if not rsi.eof then
	   send_mail="Y"
	   branchname=trim(rsi("branchname"))
	   body="致：各位主管暨同仁<br><br>"
	   body=body & "提供 貴單位各營洽代理人請款尚未對應交辦收費已超過管制期限之案件，明細如下，敬請至「規費提列不足銷管作業」儘速完成銷管，謝謝。" & "<br><br>"
	   body=body & "●尚未對應交辦收費之代收明細：<font color='red'>"
	   if ptodo="2" then
			body=body & "管制期限：" & dateadd("d",+1,pdate) & "~" & dateadd("d",-1,date()) & "(已超過1~2天)<br><br>" 
			sub_title="(期限已超過1~2天)"
	   elseif ptodo="3" then
			body=body & "管制期限：~" & pdate & "(已超過3天)<br><br>" 
			sub_title="(期限已超過3天)"	   
	   end if
	   body=body & "</font>"
	   
	   while not rsi.eof
	      scode=trim(rsi("scode1"))
	      cnt=rsi("cnt")
	      sc_name=trim(rsi("sc_name"))
	      
	      StrToList1=StrToList1 & scode & "@saint-island.com.tw;"  '收件者
		  body=body & "※"& sc_name & "，共" & cnt & "件<br>" 	      
		  
	      rsi.movenext
	   wend    
	end if
	rsi.close
    body=body & "<br>"

    oldscode = ""
	sql="select c.branch,a.scode1,a.seq,a.seq1,a.country,a.cappl_name,a.tp_no,a.tp_no1,c.step_date,c.rs_detail "
	sql = sql & " ,(select sc_name from sysctrl.dbo.scode where scode=a.scode1) as sc_name "
	sql = sql & " ,(select branchname from sysctrl.dbo.branch_code where branch=c.branch) as branchname "
	sql = sql & " ,(select (dn_money*dn_rate)+pos_fee+hand_fee from exch_temp where seq=c.seq and seq1=c.seq1 and exch_no=c.exch_no) as dnnt_money "
    sql = sql & " from todo_exp t inner join exp as a on a.seq=t.seq and a.seq1=t.seq1 " 
    sql = sql & " inner join step_exp c on c.seq=t.seq and c.seq1=t.seq1 and c.rs_sqlno = t.rs_sqlno"
	sql = sql & " where t.syscode='" & session("syscode") & "'"	
	sql = sql & " and t.dowhat in ('SC_TRF') and t.job_status in ('NN','NX')"
	if ptodo="2" then
	   sql = sql & " and dateadd(month,1,c.step_date)>'" & pdate & "' and dateadd(month,1,c.step_date)<'" & date() & "'"
	elseif ptodo="3" then
	   sql = sql & " and dateadd(month,1,c.step_date)<='" & pdate & "'"
	end if   
	sql = sql & " order by c.branch,convert(int,substring(a.scode1,2,5)),c.step_date "
	'Response.Write sql&"<br>"
	'Response.End
	rsi.open sql,conn,1,1
	if not rsi.eof then
        body=body & "<table width='90%' border='1' cellspacing='0' cellpadding='0' style='font-size:10pt'>"
        while not rsi.eof
            if oldscode <> rsi("scode1") then
	            body=body & "<tr align='center' style='BACKGROUND-COLOR:#CCFFFF'>"
		        body=body & "<td nowrap>營洽</td><td nowrap>本所編號</td><td nowrap>案件名稱</td><td nowrap>國外所案號</td><td nowrap>區所收文日</td>"
		        body=body & "<td nowrap>收文內容</td><td nowrap>管制期限</td><td nowrap>代理人請款金額(NTD)</td>"
		        body=body & "</tr>"
            end if
            fseq = formatseq2(rsi("branch"),session("dept"),"E",rsi("seq"),rsi("seq1"),rsi("country"))
            fexpseq = formatseq2("",session("dept"),"E",rsi("tp_no"),rsi("tp_no1"),"")
            body=body & "<tr align='center'>"
            body=body & "<td nowrap>"& rsi("sc_name") &"</td>"
            body=body & "<td nowrap>"& fseq &"</td>"
            body=body & "<td nowrap>"& fMid(rsi("cappl_name"),20) &"</td>"
            body=body & "<td nowrap>"& fexpseq &"</td>"
            body=body & "<td nowrap>"& rsi("step_date") &"</td>"
            body=body & "<td nowrap>"& rsi("rs_detail") &"</td>"
            body=body & "<td nowrap>"& dateadd("m",1,rsi("step_date")) &"</td>"
            body=body & "<td nowrap>"& FormatNumber(rsi("dnnt_money"),0) &"</td>"
            body=body & "</tr>"
            
            oldscode = rsi("scode1")            
            rsi.movenext
        wend
        body=body & "</table>"
    end if
    rsi.close
    
	'body=body & "<br>◎請至" & branchname & "專利案件管理系統－＞出口營洽－＞出口案規費提列不足銷管作業或營洽工作清單之規費不足尚未銷管　查詢相關資料"
	subject = branchname & "專利案件管理系統－規費不足管制期限稽催通知" & sub_title
	
	
	'抓取部門主管、區所主管
	fsql="select master_scode from sysctrl.dbo.grpid where grpclass='" & session("se_branch") & "' and (grpid='000' or grpid='P000') order by grplevel desc "
	rsi.open fsql,conn,1,1
	StrToList=""
	while not rsi.eof
	   StrToList=StrToList & trim(rsi("master_scode")) & "@saint-island.com.tw;"  '收件者	   
	   rsi.movenext
	wend
	rsi.close
	
	
	'抓取執委
	if ptodo="3" then
		fsql="select scode from sysctrl.dbo.scode_roles where dept='P' and syscode='" & session("se_branch") & "BRP' and roles='chair' "
		rsi.open fsql,conn,1,1
		if not rsi.eof then
			StrToList=StrToList & trim(rsi("scode")) & "@saint-island.com.tw;"  '收件者	   
		end if   
		rsi.close
	end if
	'Response.Write "正式收件者：" & StrToList & "<br>"
	'Response.Write "正式副本：" & StrToList1 & "<br>"
	BCCStrToList = ""
	Select Case Request.ServerVariables("SERVER_NAME")
		Case "web02"
			body=body&"<br>正式收件者：" & StrToList & "<br>"
			body=body&"正式副本：" & StrToList1 & "<br>"
		
			'StrToList="m983@saint-island.com.tw"  '收件者
			StrToList=session("scode")&"@saint-island.com.tw"  '收件者
			'Sender="m983" '寄件者
			Sender=session("scode")
			'Sender="administrator"	'寄件者
			StrToList1=""	'副本
			subject=subject & "("&Request.ServerVariables("SERVER_NAME")&"測試信)"
	    
		Case "web01"
			StrToList=session("scode")&"@saint-island.com.tw"  '收件者
			'Sender=session("scode") '寄件者
			Sender="administrator"	'寄件者
			StrToList1=""	'副本
			subject=subject & "("&Request.ServerVariables("SERVER_NAME")&"測試信)"
		Case else
			'Sender=session("scode") '寄件者
			Sender="administrator"	'寄件者
			'BCCStrToList="m983@saint-island.com.tw"
	End Select
	
	'Response.Write "body=" & body & "<br>"
	'Response.Write "subject=" & subject & "<br>"
	'Response.End
	
	Sendmail=send_mail
End Function




Sub DoSendMail(S,B)  '******發信開始
	i=0
	Reciper = StrToList
	body_temp= B 
	set objMail=createobject("cdonts.newmail")
	objmail.mailformat = cdomailformatmime
	objmail.from= Sender&"@saint-island.com.tw"
	objmail.to =  Reciper
	objmail.cc =  StrToList1
	if BCCStrToList<>empty then
	    objmail.bcc = BCCStrToList
	end if
	objmail.mailformat=cdomailformatMIME
	objmail.BodyFormat=cdoBodyFormatHTML
	objmail.subject= S
	objmail.body= body_temp
	objmail.send
	set objmail=nothing
End Sub  '******發信結束


function formatseq2(branch,dept,pe,seq,seq1,country)
    lseq = ""
    if branch<>empty then
	    lseq = lseq & branch
	end if
	lseq = lseq & ucase(dept)
	if pe="E" then
		lseq = lseq & ucase(pe)
	end if
	lseq = lseq & seq
	if seq1 <> "_" then
		lseq = lseq & "-" & seq1
	end if
	if country <> empty then
		IF country<>"T" then
			lseq = lseq & "-" & ucase(country)
		End IF
	end if
	formatseq2 = lseq
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

	
%>
