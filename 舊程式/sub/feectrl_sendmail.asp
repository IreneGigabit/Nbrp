
<%
Begin_time=now


Dim StrToList,StrToList1, Sender,subject,body,BCCStrToList

function feectrl_onclick()
   mail_flag="N"
   
   '�ˬdcust_code.code_type=Z and cust_code.cust_code=PEfc_mail_date�����
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
        '�����W�L1~2��
		task=sendmail("2")
		if task="Y" then
		    DoSendMail subject,body
		end if  
		'�����W�L3��
		task1=sendmail("3")
		if task1="Y" then
		    DoSendMail subject,body
		end if   
		'�ק�cust_code��Email�q�����
		if task="Y" or task1="Y" then
		    usql="update cust_code set form_name='" & date() & "' where code_type='Z' and cust_code='PEfc_mail_date' "
		    conn.execute usql
		end if 
   end if		
end function
'ptodo="2"�����W�L1-2�ѡA�q����ϩҥD��,ptodo="3"�����W�L3�ѡA�q������e
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
	StrToList1=""	'�ƥ�
	send_mail="N"
	if not rsi.eof then
	   send_mail="Y"
	   branchname=trim(rsi("branchname"))
	   body="�P�G�U��D�޺[�P��<br><br>"
	   body=body & "���� �Q���U�笢�N�z�H�дک|��������즬�O�w�W�L�ި�������ץ�A���Ӧp�U�A�q�Цܡu�W�O���C�����P�ާ@�~�v���t�����P�ޡA���¡C" & "<br><br>"
	   body=body & "���|��������즬�O���N�����ӡG<font color='red'>"
	   if ptodo="2" then
			body=body & "�ި�����G" & dateadd("d",+1,pdate) & "~" & dateadd("d",-1,date()) & "(�w�W�L1~2��)<br><br>" 
			sub_title="(�����w�W�L1~2��)"
	   elseif ptodo="3" then
			body=body & "�ި�����G~" & pdate & "(�w�W�L3��)<br><br>" 
			sub_title="(�����w�W�L3��)"	   
	   end if
	   body=body & "</font>"
	   
	   while not rsi.eof
	      scode=trim(rsi("scode1"))
	      cnt=rsi("cnt")
	      sc_name=trim(rsi("sc_name"))
	      
	      StrToList1=StrToList1 & scode & "@saint-island.com.tw;"  '�����
		  body=body & "��"& sc_name & "�A�@" & cnt & "��<br>" 	      
		  
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
		        body=body & "<td nowrap>�笢</td><td nowrap>���ҽs��</td><td nowrap>�ץ�W��</td><td nowrap>��~�Ү׸�</td><td nowrap>�ϩҦ����</td>"
		        body=body & "<td nowrap>���夺�e</td><td nowrap>�ި����</td><td nowrap>�N�z�H�дڪ��B(NTD)</td>"
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
    
	'body=body & "<br>���Ц�" & branchname & "�M�Q�ץ�޲z�t�ΡС֥X�f�笢�С֥X�f�׳W�O���C�����P�ާ@�~���笢�u�@�M�椧�W�O�����|���P�ޡ@�d�߬������"
	subject = branchname & "�M�Q�ץ�޲z�t�ΡгW�O�����ި�����]�ʳq��" & sub_title
	
	
	'��������D�ޡB�ϩҥD��
	fsql="select master_scode from sysctrl.dbo.grpid where grpclass='" & session("se_branch") & "' and (grpid='000' or grpid='P000') order by grplevel desc "
	rsi.open fsql,conn,1,1
	StrToList=""
	while not rsi.eof
	   StrToList=StrToList & trim(rsi("master_scode")) & "@saint-island.com.tw;"  '�����	   
	   rsi.movenext
	wend
	rsi.close
	
	
	'������e
	if ptodo="3" then
		fsql="select scode from sysctrl.dbo.scode_roles where dept='P' and syscode='" & session("se_branch") & "BRP' and roles='chair' "
		rsi.open fsql,conn,1,1
		if not rsi.eof then
			StrToList=StrToList & trim(rsi("scode")) & "@saint-island.com.tw;"  '�����	   
		end if   
		rsi.close
	end if
	'Response.Write "��������̡G" & StrToList & "<br>"
	'Response.Write "�����ƥ��G" & StrToList1 & "<br>"
	BCCStrToList = ""
	Select Case Request.ServerVariables("SERVER_NAME")
		Case "web02"
			body=body&"<br>��������̡G" & StrToList & "<br>"
			body=body&"�����ƥ��G" & StrToList1 & "<br>"
		
			'StrToList="m983@saint-island.com.tw"  '�����
			StrToList=session("scode")&"@saint-island.com.tw"  '�����
			'Sender="m983" '�H���
			Sender=session("scode")
			'Sender="administrator"	'�H���
			StrToList1=""	'�ƥ�
			subject=subject & "("&Request.ServerVariables("SERVER_NAME")&"���իH)"
	    
		Case "web01"
			StrToList=session("scode")&"@saint-island.com.tw"  '�����
			'Sender=session("scode") '�H���
			Sender="administrator"	'�H���
			StrToList1=""	'�ƥ�
			subject=subject & "("&Request.ServerVariables("SERVER_NAME")&"���իH)"
		Case else
			'Sender=session("scode") '�H���
			Sender="administrator"	'�H���
			'BCCStrToList="m983@saint-island.com.tw"
	End Select
	
	'Response.Write "body=" & body & "<br>"
	'Response.Write "subject=" & subject & "<br>"
	'Response.End
	
	Sendmail=send_mail
End Function




Sub DoSendMail(S,B)  '******�o�H�}�l
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
End Sub  '******�o�H����


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

	
%>
