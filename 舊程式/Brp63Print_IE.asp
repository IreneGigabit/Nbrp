<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

prgid = lCase(Request("prgid"))
session("prgid") = prgid

HTProgCap="�ӿ�ӽЮѦC�L-IG�o���M�Q�ӽЮ�"
HTProgCode="brp63"
HTProgPrefix="brp63"
%>
<!--#include file="../inc/server.inc" -->
<!--#include file="../sub/Server_cbx.vbs" -->
<!--#include file="../sub/Server_conn.vbs" -->
<!--#include file="../sub/Server_conn_unicode.vbs" -->
<!--#include file="../brp6m/Brp63Print_sub_IG_E.asp"-->
<!--#include file="../brp6m/brpform/brp63_IE_form_1.asp" -->
<script language="vbscript">
window.parent.tt.rows="50%,50%"
</script>
<%
	DIM rs 
	dim RSreg, sql
	cust_area = request("cust_area")
	cust_seq = request("cust_seq")
	in_scode = request("in_scode")
	in_no = request("in_no")
	prgid = request("prgid")
	arcase=request("arcase")
	branch=trim(session("se_branch"))
	Set rs = CreateObject("ADODB.RecordSet")
	Set rs1 = CreateObject("ADODB.RecordSet")
	Set rs2 = CreateObject("ADODB.RecordSet")
	
	
    SQL = "select * from vdmpall where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
	'Response.Write SQL & "<BR>"
    'response.End
    
    rs.Open SQL,conn,1,1
    
while not rs.EOF

    Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
	objStream.Charset = "utf-8"
	objStream.Position = objStream.Size
	
	
    Dim Doc_Head_1
    Dim Doc_Body_1
    'Dim Doc_Body_2
    Dim Doc_Tail_1
    
    Doc_Head_1 = DocHead_1()  
    objStream.WriteText = Doc_Head_1
    
    Doc_Body_1 = DocBody_1()
    objStream.WriteText = Doc_Body_1
    
    Space_String = SpaceString()

    '�ץ�
    Doc_Body_2 = DocBody_2()
    Doc_Body_2 = ReplaceData(Doc_Body_2, "#case_no#", "10000","empty")        
    'Response.Write "reality =" & trim(rs("reality")) & "<BR>"
    'response.End    
    '�@�֥ӽй���f�d
    IF trim(rs("reality"))= "Y" then
        Doc_Body_2 = ReplaceData(Doc_Body_2, "#reality#", ToUnicode2("�O"),"empty")
    else
        Doc_Body_2 = ReplaceData(Doc_Body_2, "#reality#", ToUnicode2("�_"),"empty")
	END IF
    '�ưȩҩΥӽФH�ץ�s��
    fseq = formatseq2(session("se_branch"),"P","",rs("seq"),rs("seq1"),"")&"-"&Trim(rs("scode1"))
    fseq_1 = formatseq2(session("se_branch"),"P","",rs("seq"),rs("seq1"),"")
    Doc_Body_2 = ReplaceData(Doc_Body_2, "#seq#", fseq,"empty")
    objStream.WriteText = Doc_Body_2
    

    '�ťզ�
    objStream.WriteText = Space_String
    
    
    Doc_Body_3 = DocBody_3()
    '����o���W�� '�^��o���W��
    Doc_Body_3 = ReplaceData(Doc_Body_3, "#cappl_name#", replace(ToUnicode2(Trim(rs("cappl_name"))),"&","&amp;"),"empty")
    Doc_Body_3 = ReplaceData(Doc_Body_3, "#eappl_name#", ToXmlUnicode(trim(rs("eappl_name"))),"empty")
    objStream.WriteText = Doc_Body_3
    '�ťզ�
    objStream.WriteText = Space_String
    
    'Call Dmp_apcust_data_Function ���� �ӽФH
        CALL Dmp_apcust_data_Function(in_scode,in_no)
    'Call Agt_data_1n2_Function ���� �N�z�H1 & �N�z�H2
        CALL Agt_data_1n2_Function(in_scode,in_no)
    'Call Ant_data_Function ���� �o���H
        CALL Ant_data_Function(in_scode,in_no,"IG_E")
     
    '�D�i�u�f���j��
    Doc_Body_6 = DocBody_6_1  
    if rs("exhibitor")="Y" then '�Ѯi�εo������J�����o�ͤ��
        exh_date = format_date_char(rs("exh_date")) 'YYYY/MM/DD
    else
        exh_date = ""
    end if
    Doc_Body_6 = ReplaceData(Doc_Body_6, "#exh_date#", exh_date,"empty")
    objStream.WriteText = Doc_Body_6
    '�ťզ�
    objStream.WriteText = Space_String

     '�D�i�u���v�j��.
'     iSQL = ""
'     iSQL = "  Select a.prior_date,a.prior_country,a.prior_no,a.prior,"
'     iSQL = iSQL & " a.case1,(select code_name from cust_code where code_type='case1' and cust_code=a.case1 and mark in ('1','2','3')) as case1nm  , a.mprior_access"
'     iSQL = iSQL & " from dmp a inner join vdmpall b on a.seq = b.seq and a.seq1 = b.seq1 "
'     iSQL = iSQL & " where b.in_scode='" & in_scode & "' and b.in_no='" & in_no & "' and a.prior = 'Y' "
'        'response.Write isql 
'        'response.End
'     rs1.Open iSQL,conn,1,1
'     
'     if not rs1.EOF then
'          for i=1 to rs1.RecordCount 
'              
'              iisql = ""
'              iisql = " select b.coun_code,b.coun_cname,b.coun_ename From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code "
'              iisql = iisql & " where a.coun_code = '"&trim(rs1("prior_country"))&"'"
'            
'            'REsponse.Write iisql
'            'response.End
'            
'            rs2.Open iisql,conn,1,1
'            if not rs2.EOF then 
'                Country_name = empty
'                Country_name =  trim(rs2("coun_code"))&trim(rs2("coun_cname"))
'            end if
'               rs2.Close
'            if trim(rs1("prior_date"))<>empty then 
'                Format_prior_date = year(trim(rs1("prior_date"))) & "/" & String(2 - Len(month(trim(rs1("prior_date")))), "0") & month(trim(rs1("prior_date"))) & "/" & String(2 - Len(day(trim(rs1("prior_date")))), "0") & day(trim(rs1("prior_date")))        
'            END IF
'            
'            if trim(rs1("prior")) = "Y" THEN
'                Doc_Body_7 = DocBody_7()
'                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_num#", i,"empty")
'                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_country#", Country_name,"empty")           
'                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_date#", Format_prior_date,"empty")
'                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_no#", trim(rs1("prior_no")),"empty")          
'                objStream.WriteText = Doc_Body_7
'                
'                if trim(rs1("prior_country")) = "JA" THEN 
'                
'                    Doc_Body_7_1 = DocBody_7_1()
'                    Doc_Body_7_1 = ReplaceData(Doc_Body_7_1, "#case1nm#", trim(rs1("case1nm")),"empty")             
'                    Doc_Body_7_1 = ReplaceData(Doc_Body_7_1, "#mprior_access#", trim(rs1("mprior_access")),"empty")       
'                    objStream.WriteText = Doc_Body_7_1                            
'                END IF    
'                
'            END IF
'            
'           
'
'            rs1.MoveNext 
'            '�ťզ�
'          objStream.WriteText = Space_String
'          NEXT
'     END IF
'     rs1.Close

     isql = "SELECT a.prior_yn, a.prior_no, a.prior_country, a.prior_date, a.mprior_access, a.prior_case1"
     isql = isql & ", c.coun_code, c.coun_cname, c.coun_ename"   
     isql = isql & ", (SELECT mark1 FROM cust_code WHERE code_type = 'case1' AND cust_code = a.prior_case1) AS case1nm_T"
     isql = isql & ", (SELECT code_name FROM cust_code WHERE code_type = 'pecase1' AND cust_code = a.prior_case1) AS case1nm"
     isql = isql & " FROM dmp_prior AS a"
     isql = isql & " INNER JOIN vdmpall AS b ON a.seq = b.seq AND a.seq1 = b.seq1"
     'isql = isql & " INNER JOIN vdmpall AS b ON a.dmp_sqlno = b.dmp_sqlno "
     isql = isql & " LEFT JOIN sysctrl.dbo.IPO_country AS c ON a.prior_country = c.ref_coun_code"
     isql = isql & " WHERE b.in_scode = '" & in_scode & "'"
     isql = isql & " AND b.in_no = '" & in_no & "'"
     isql = isql & " AND a.prior_yn = 'Y'"
    'response.Write isql 
    'response.End
     rs1.Open isql,conn,1,1
     If Not rs1.EOF Then
          For i=1 To rs1.RecordCount             
                country_name =  Trim(rs1("coun_code")) & Trim(rs1("coun_cname"))
                
                If trim(rs1("prior_date"))<>Empty Then 
                    format_prior_date = year(trim(rs1("prior_date"))) & "/" & String(2 - Len(month(trim(rs1("prior_date")))), "0") & month(trim(rs1("prior_date"))) & "/" & String(2 - Len(day(trim(rs1("prior_date")))), "0") & day(trim(rs1("prior_date")))        
                End If 
                              
                Doc_Body_7 = DocBody_7()
                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_num#", i, "empty")
                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_country#", country_name, "empty")           
                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_date#", format_prior_date, "empty")
                Doc_Body_7 = ReplaceData(Doc_Body_7, "#prior_no#", Trim(rs1("prior_no")), "empty")      
                    
                objStream.WriteText = Doc_Body_7
                
                Select Case Trim(rs1("prior_country"))
                    Case "JA"             
                        Doc_Body_7_1 = DocBody_7_1()
                        Doc_Body_7_1 = ReplaceData(Doc_Body_7_1, "#case1nm#", Trim(rs1("case1nm")),"empty")             
                        Doc_Body_7_1 = ReplaceData(Doc_Body_7_1, "#mprior_access#", Trim(rs1("mprior_access")), "empty")      
                         
                        objStream.WriteText = Doc_Body_7_1
                    Case "KO"
                        Doc_Body_7_2 = DocBody_7_2()
                        Doc_Body_7_2 = ReplaceData(Doc_Body_7_2, "#mprior_access#", "�洫", "empty")      
                         
                        objStream.WriteText = Doc_Body_7_2                                                      
                End Select    
                
                '�ťզ�
                objStream.WriteText = Space_String                    
            
                rs1.MoveNext 
          Next
     End If
     rs1.Close
     
    '�D�i�Q�Υͪ�����
    Doc_Body_7_2 = DocBody_7_2
    objStream.WriteText = Doc_Body_7_2
    '�ťզ�
    objStream.WriteText = Space_String
    
    '�ͪ����Ƥ����H�s
    Doc_Body_8 = DocBody_8()
    objStream.WriteText = Doc_Body_8
    '�ťզ�
    objStream.WriteText = Space_String
     
    '�n�����H�N�ۦP�Ч@�b�ӽХ��o���M�Q���P��-�t�ӽзs���M�Q
    Doc_Body_81 = DocBody_81()
    if trim(rs("same_apply"))="Y" then
        Doc_Body_81 = ReplaceData(Doc_Body_81, "#same_apply#", "�O","empty") 
    else
        Doc_Body_81 = ReplaceData(Doc_Body_81, "#same_apply#", "","empty")
    end if
    objStream.WriteText = Doc_Body_81
    '�ťզ�
    objStream.WriteText = Space_String

    '���奻��T ,�~�奻��T ,ú�O��T
    Doc_Body_9 = DocBody_9()
    objStream.WriteText = Doc_Body_9
	 '20170524 �W�[���ک��Y�ﶵ
	 Title_string= Dmp_rectitle_Function(in_scode,in_no,request("receipt_title"))
	 objStream.WriteText = ReplaceData(Dmp_receipt_title(), "#rectitle_name#", replace(ToUnicode2(Title_string),"��","&amp;"),"empty")
    '�ťզ�
    objStream.WriteText = Space_String
    
    '���e�ѥ�
    Doc_Body_10 = DocBody_10()    
    Doc_Body_10 = ReplaceData(Doc_Body_10, "#seq#", formatseq2(session("se_branch"),"P","",rs("seq"),rs("seq1"),""),"empty")
    objStream.WriteText = Doc_Body_10
    '�ťզ�
    objStream.WriteText = Space_String

    '�򥻸�� ;�ӤH��ƶ}�Y
    Doc_Body_11 = DocBody_11()    
    objStream.WriteText = Doc_Body_11

    'Call Dmp_apcust_data_1_function ���� �򥻸�ƪ�-�ӽФH
    CALL Dmp_apcust_data_1_function(in_scode,in_no)
    'Call Agt_data_1_Function ���� �򥻸�ƪ�-�N�z�H1
    CALL Agt_data_1_Function(in_scode,in_no)
    'Call Agt_data_2_Function ���� �򥻸�ƪ�-�N�z�H2
    CALL Agt_data_2_Function(in_scode,in_no)
    'Call Ant_data_1_Function ���� �򥻸�ƪ�-�o���H
    CALL Ant_data_1_Function(in_scode,in_no,"IG_E")

    Doc_Tail_1 = DocTail_1()
    objStream.WriteText = Doc_Tail_1


   ' response.Write "tattach_name="& tattach_name & "<br>"
	'response.Write tattach_path & "\" & tattach_name & "<BR>"
	'response.End 
    filepath =fseq_1&"-�o��.doc"
    'response.Write err.number &"("& err.Description  &")--4<Br>"
	objStream.SaveToFile "D:\Inetpub\wwwroot\brp\reportdata\"&filepath, 2
	filesize = objStream.Size
	'Response.Write "size="& filesize & "<BR>"
	
	objStream.Close
    Set objStream = Nothing		
    
	rs.MoveNext 
wend
rs.Close 
%>

<script type="text/vbscript" language="vbscript">
	<%if ERR.number=0 then
		sql = "update case_dmp set new='P'+substring(NEW,2,50) "
		sql = sql & ",receipt_title='" & request("receipt_title") & "' "
		sql = sql & ",rectitle_name='" & Title_string & "' "
		sql = sql & "where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
		conn.Execute(sql)
	%>
		'window.open "ReportWord\" & "<%=session("se_scode")%>" & "�o���M�Q�ӽЮ�.doc"
		tfile = "?"& CStr(Year(Now())) & Right("0" & CStr(Month(Now())), 2) & Right("0" & CStr(Day(Now())), 2) & Right("0" & CStr(Hour(Now())), 2) & Right("0" & CStr(Minute(Now())), 2) & Right("0" & CStr(Second(Now())), 2)
		window.open "../reportdata/<%=filepath%>"&tfile,"myWindowOne", "width=1270 height=830 top=0 left=0 toolbar=yes menubar=yes resizable=yes scrollbars=yes "
	<%else%>
		msgbox "�o���M�Q�ӽЮ� Word ���ͥ���!!!"
	<%end if%>
	
window.parent.Eblank.location.href = "../brp2m/Brp22ChoP.asp?prgid=" & "<%=prgid%>" & "&in_scode=" & "<%=request("in_scode")%>" & "&in_no=" & "<%=request("in_no")%>" & "&cust_area=" & "<%=request("cust_area")%>" & "&cust_seq=" & "<%=request("cust_seq")%>" & "&arcase=" & "<%=request("arcase")%>"


	
	
</script>

<%
set rs = nothing
set rs1 = nothing
Response.Write "now2=" & now & "<br>"
%>
