<%
'轉換工作項目名稱
function work_code(tdept,tcode1,tcode2)
  if len(trim(tcode1)) > 1 then
     tcode1=left(trim(tcode1),1)
  else
     tcode1=trim(tcode1)
  end if
  strSQLf_c = "select code_name from work_code where work_dept='" & trim(tdept) _
             & "' and class = '" & tcode1 & "' and code = '" & trim(tcode2) & "'"
             
  'Response.Write strSQLf_c
  Set connf_c = Server.CreateObject("ADODB.Connection")
      connf_c.Open session("sin03fdbs")
  set rsf_c=connf_c.Execute(strSQLf_c)
  if rsf_c.eof then
     work_code=trim(a)
  else
     work_code=trim(rsf_c("code_name"))
  end if
  rsf_c.close
  set rsf_c=nothing
  connf_c.close
  set connf_c=nothing
end function

'秀工作項目選擇項
'a=class
'b=code
'c=己選項目 
function work_code02(a,b,c)
  strSQLf_c = "select code,code_name from fdbs_code where class = '" _
  & trim(a) & "' and code like '" & trim(b) & "%' order by sql"
 
  Set connf_c = Server.CreateObject("ADODB.Connection")
      connf_c.Open session("sin03fdbs")
  response.write "<option value='' style='color:blue'>請選擇" & chr(13)
  set rsf_c=connf_c.Execute(strSQLf_c)
  while not rsf_c.eof
        if trim(rsf_c("code")) = trim(c) then
           response.write "<option value='"&trim(rsf_c("code"))&"' selected>" &trim(rsf_c("code_name")) & chr(13)
        else
           response.write "<option value='"&trim(rsf_c("code"))&"'>" &trim(rsf_c("code_name")) & chr(13)
        end if     
     rsf_c.movenext
  wend   
  rsf_c.close
  set rsf_c=nothing
  connf_c.close
  set connf_c=nothing
end function

'轉換承辦人員
function work_code01(a)
  strSQLf_c = "select scode,sc_name from scode where scode='"& trim(a) &"'"
  'Response.Write strSQLf_c
  Set connf_c = Server.CreateObject("ADODB.Connection")
      connf_c.Open session("ODBCDSN")
  set rsf_c=connf_c.Execute(strSQLf_c)
  if rsf_c.eof then
     work_code01=trim(a)
  else
     work_code01=trim(rsf_c("sc_name"))
  end if
  rsf_c.close
  set rsf_c=nothing
  connf_c.close
  set connf_c=nothing
end function

'轉換處理狀態
function stat_code(a)
  strSQLf_c = "select code_name from fdbs_code where " _
             & " class = 'status' and code = '" & trim(a) & "'"
  'Response.Write strSQLf_c
  Set connf_c = Server.CreateObject("ADODB.Connection")
      connf_c.Open session("sin03fdbs")
  set rsf_c=connf_c.Execute(strSQLf_c)
  if rsf_c.eof then
     stat_code=trim(a)
  else
     stat_code=trim(rsf_c("code_name"))
  end if
  if isnull(trim(a)) then
     stat_code="未分案"
  end if
  rsf_c.close
  set rsf_c=nothing
  connf_c.close
  set connf_c=nothing
end function

'轉換人員薪號->姓名
sub tran_name(a)
  strSQLt = "select sc_name from scode where scode='"&trim(a)&"'"
  Set connt = Server.CreateObject("ADODB.Connection")
      connt.Open session("ODBCDSN")
  set rst=connt.Execute(strSQLt)
  if rst.eof then
      response.write trim(a)
  else
      Response.Write rst("sc_name")
  end if
  rst.close
  set rst=nothing
  connt.close
  set connt=nothing
end sub

''轉換人員薪號->所屬部門
sub tran_dept(a)
  strSQLt = "select grpname from grpid " _
          & "where grpclass='TPI' and grptype='N' " _
          & "and grpid=(select grpid from scode_group where grpclass='TPI' and scode='"&trim(a)&"' and grptype='N') "
  Set connt = Server.CreateObject("ADODB.Connection")
      connt.Open session("ODBCDSN")
  set rst=connt.Execute(strSQLt)
  if rst.eof then
      response.write trim(a)
  else
      Response.Write trim(rst("grpname"))
  end if
  rst.close
  set rst=nothing
  connt.close
  set connt=nothing
end sub

'轉換組別->姓名
sub tran_team(a)
  strSQLt = "select grpname from grpid where grpclass='TPI' and grpid='"&trim(a)&"' and grptype='N'"
  Set connt = Server.CreateObject("ADODB.Connection")
      connt.Open session("ODBCDSN")
  set rst=connt.Execute(strSQLt)
  if rst.eof then
      response.write trim(a)
  else
      Response.Write rst("grpname")
  end if
  rst.close
  set rst=nothing
  connt.close
  set connt=nothing
end sub

'轉換作業類別->姓名
sub tran_work1(tdept,a)
  strSQLt = "select code_name from work_code where work_dept = '" & tdept _
            & "' and class='" & trim(a) & "' and code = '____'"
  Set connt = Server.CreateObject("ADODB.Connection")
      connt.Open session("sin03fdbs")
  set rst=connt.Execute(strSQLt)
  if rst.eof then
      response.write trim(a)
  else
      Response.Write rst("code_name")
  end if
  rst.close
  set rst=nothing
  connt.close
  set connt=nothing
end sub

'由薪號判斷所屬的Grpid
function s_g0(a)
  strSQLg = "select scode,grpid from scode_group where scode='"& trim(a) &"' and grpclass='TPI' and grptype='N'"
  Set conng = Server.CreateObject("ADODB.Connection")
      conng.Open session("ODBCDSN")
  set rsg=conng.Execute(strSQLg)
  if rsg.eof then
      s_g0=trim(a)
  else
      s_g0=trim(rsg("grpid"))
  end if
  rsg.close
  set rsg=nothing
  conng.close
  set conng=nothing
end function

'show個組or三組承辦人員
'a=薪號(抓所屬個組所有承辦人員)
'b=0/1(0:表不加全部;1:表加全部)-增加<option>
'c=判斷已選擇項(表選擇的<option>)
sub pr_scode(a,b,c)
  Set connchk = Server.CreateObject("ADODB.Connection")
  connchk.Open session("ODBCDSN")
  strSQLchk = "SELECT s.scode,s.sc_name FROM scode s right JOIN scode_group g "_
         & "ON s.scode = g.scode where g.grpid=(select grpid from scode_group where scode='"_
         & trim(a) &"' and grpclass='TPI' and grptype='N') order by s.sc_name"
  'response.write strSQLchk
  set rschk = connchk.Execute(strSQLchk)
  response.write "<select name='pr_scode'>" & chr(13)
  response.write "<option value='' style='color:blue'>請選擇" & chr(13)
  if b=1 then 
     if c="%" then
        response.write "<option value='%' selected>全部" & chr(13)
     else
        response.write "<option value='%'>全部"& chr(13)
     end if   
  end if
  While Not rschk.EOF
  if b=1 then
     if trim(rschk("scode")) = trim(c) then
        response.write "<option value='"& rtrim(rschk("scode")) & "' selected>" & rtrim(rschk("sc_name")) & "</option>"  & chr(13)
     else           
        response.write "<option value='"& rtrim(rschk("scode")) & "'>" & rtrim(rschk("sc_name")) & "</option>"  & chr(13)
     end if        
   else
        response.write "<option value='"& rtrim(rschk("scode")) & "'>" & rtrim(rschk("sc_name")) & "</option>"  & chr(13)
   end if     
        rschk.MoveNext
  WEnd
  response.write "</select>"
  rschk.close
  connchk.close  
  set connchk = nothing 
end sub

'支援人力
'a=薪號
sub spr_scode(a,b)
  Set conns = Server.CreateObject("ADODB.Connection")
  conns.Open session("ODBCDSN")
  strSQLs = "SELECT s.scode,s.sc_name FROM scode s right JOIN scode_group g ON s.scode = g.scode "_
          & "where g.grpid=(substring((select grpid from scode_group where scode='t234' and grpclass='TPI' and grptype='N'),1,1)+'00Z') order by s.sc_name"
  'response.write strSQLs
  set rss = conns.Execute(strSQLs)
  response.write "<select name='spr_scode'>" & chr(13)
  response.write "<option value='' style='color:blue'>請選擇" & chr(13)
  While Not rss.EOF
      if trim(rss("scode")) = trim(b) then
        response.write "<option value='"& rtrim(rss("scode")) & "' selected>" & rtrim(rss("sc_name")) & "</option>"  & chr(13)
      else
        response.write "<option value='"& rtrim(rss("scode")) & "'>" & rtrim(rss("sc_name")) & "</option>"  & chr(13)
      end if  
        rss.MoveNext
  WEnd
  response.write "</select>"
  rss.close
  conns.close  
  set conns = nothing 
end sub

'轉換國別代碼
function coun(aaa)
    Set connc = Server.CreateObject("ADODB.Connection")
    connc.Open session("sin03mag")
	strSQLc = "SELECT coun_code,coun_c FROM coun where coun_code='" & trim(aaa) &"'"
    set rsc = connc.Execute(strSQLc)
    if rsc.eof then
       coun=aaa
    else
       coun=trim(rsc(1))
    end if
    rsc.close
    connc.close
    set connc=nothing
end function
'show work_code1
'sub get_work1(tdept)
'	Set connchk = Server.CreateObject("ADODB.Connection")
'    connchk.Open session("sin03fdbs")
'    strsql="select class,code_name from work_code " _
'           & "where work_dept='" & tdept & "' and code = '____' " _
'           & "order by sql,class"
'    set wrs=connchk.execute(strsql)
'    while not wrs.eof
'		cls=trim(wrs("class"))
'		codename=trim(wrs("code_name"))
'		tcls=len(cls)
'		if tcls = 1 then
'		   response.write "<input type='radio' name='work_code1' id='work_code1' value="&cls&" onclick='"&work_code02(tdept,cls)&"'>"
'		else
'		   response.write "<input type='radio' name='work_code1 id='work_code1' value="&cls&">"&codename&"(無指示)"		   
'		end if
'		
'		wrs.movenext
'	wend
'end sub
%>