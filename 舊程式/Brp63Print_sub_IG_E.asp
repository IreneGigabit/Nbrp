<%
function format_date_char(pdate)
    if pdate<>empty then
        format_date_char = year(pdate) &"/"& string(2-len(month(pdate)),"0")&month(pdate) &"/"& string(2-len(day(pdate)),"0")&day(pdate) 
    end if
end function
'�ӽЮѤ� . �ӽФH�j��
Function Dmp_apcust_data_Function(in_scode,in_no)
       '�ӽФH�j��.
     iSQL = ""
     iSQL = "select b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename"
     iSQL = iSQL & " from dmp_apcust a,apcust b "
     iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'  and kind='A' and a.apsqlno=b.apsqlno"
        'response.Write isql 
        'response.End
     rs1.Open iSQL,conn,1,1
     
     if not rs1.EOF then
          for i=1 to rs1.RecordCount 
            Dmp_apcust_data_local = Dmp_apcust_data()
            iisql = ""
            iisql = " select b.coun_code,b.coun_cname,b.coun_ename From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code "
            iisql = iisql & " where a.coun_code = '"&trim(rs1("ap_country"))&"'"
            
            'REsponse.Write iisql
            'response.End
            
            rs2.Open iisql,conn,1,1
            if not rs2.EOF then 
                Country_name = empty
                Country_name =  trim(rs2("coun_code"))&trim(rs2("coun_cname"))
            end if
               rs2.Close
            
            Dmp_apcust_data_local = ReplaceData(Dmp_apcust_data_local, "#apply_num#", i,"empty")
            Dmp_apcust_data_local = ReplaceData(Dmp_apcust_data_local, "#ap_country#", Country_name,"empty")
            
            '���ꤽ�q
            if Mid(trim(rs1("apclass")),1,1) = "A" Then
                Title_cname = "����W��"
                Title_ename = "�^��W��"
                Cname_string = trim(rs1("ap_cname1"))&trim(rs1("ap_cname2")) 
                Ename_string = trim(rs1("ap_ename1"))&trim(rs1("ap_ename2")) 
            end if
            '����۵M�H
            if Mid(trim(rs1("apclass")),1,1) = "B" Then
                Title_cname = "����m�W"
                Title_ename = "�^��m�W"
                Cname_string = trim(rs1("ap_fcname"))&","&trim(rs1("ap_lcname")) 
                if Cname_string = "," then 
                    Cname_string =  trim(rs1("ap_cname1"))&trim(rs1("ap_cname2"))
                end if
                Ename_string = trim(rs1("ap_fename"))&","&trim(rs1("ap_lename")) 
                if Ename_string = "," then 
                    Ename_string =  trim(rs1("ap_ename1"))&trim(rs1("ap_ename2"))
                end if
            end if
			'20161206�~��H/���q�W�[�P�_-�Y����g�ӽФH�m&�W,�h��ܩm�W
            if Mid(trim(rs1("apclass")),1,1) = "C" Then
                Cname_string = trim(rs1("ap_fcname"))&","&trim(rs1("ap_lcname")) 
                if Cname_string = "," then 
					Title_cname = "����W��"
					Title_ename = "�^��W��"
                    Cname_string =  trim(rs1("ap_cname1"))&trim(rs1("ap_cname2"))
				else
					Title_cname = "����m�W"
					Title_ename = "�^��m�W"
                end if
                Ename_string = trim(rs1("ap_fename"))&","&trim(rs1("ap_lename")) 
                if Ename_string = "," then 
                    Ename_string =  trim(rs1("ap_ename1"))&trim(rs1("ap_ename2"))
                end if
            end if
            
            
            Dmp_apcust_data_local = ReplaceData(Dmp_apcust_data_local, "#ap_cname1_title#", Title_cname,"empty")
            Dmp_apcust_data_local = ReplaceData(Dmp_apcust_data_local, "#ap_ename1_title#", Title_ename,"empty")                    
            Dmp_apcust_data_local = ReplaceData(Dmp_apcust_data_local, "#ap_cname1#", ToUnicode2(Cname_string),"empty")
            Dmp_apcust_data_local = ReplaceData(Dmp_apcust_data_local, "#ap_ename1#", ToXmlUnicode(Ename_string),"empty")
            objStream.WriteText = Dmp_apcust_data_local
            
           
            rs1.MoveNext 
             '�ťզ�
            objStream.WriteText = Space_String
         
          Next
     end if
     rs1.Close
End function 


'�ӽЮѤ� �N�z�H1&2
Function Agt_data_1n2_Function(in_scode,in_no)
    ' 201609�P����T�{�A�אּ�Ӧ���줧�N�z�H
     iSQL = ""
     iSQL = " Select b.agt_name1,b.agt_name2 from dmp a "
     iSQL = iSQL & " inner join vdmpall c on a.dmp_sqlno = c.dmp_sqlno" 
     iSQL = iSQL & " inner join agt b on c.nagt_no = b.agt_no"
     iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
        'response.Write isql 
        'response.End
     rs1.Open iSQL,conn,1,1  
      '�N�z�H1           
       Agt_data_1_local = Agt_data_1()       
       Agt_data_1_local   = ReplaceData(Agt_data_1_local,   "#agt_name1#", Mid(trim(rs1("agt_name1")),1,1)&","&mid(trim(rs1("agt_name1")),2,2),"empty")
       objStream.WriteText = Agt_data_1_local
       '�ťզ�
          objStream.WriteText = Space_String
      '�N�z�H2
       Agt_data_2_local = Agt_data_2()  
       Agt_data_2_local = ReplaceData(Agt_data_2_local, "#agt_name2#", Mid(trim(rs1("agt_name2")),1,1)&","&mid(trim(rs1("agt_name2")),2,2),"empty")    
       objStream.WriteText = Agt_data_2_local     
         '�ťզ�
          objStream.WriteText = Space_String
     rs1.Close
END Function

'�ӽЮѤ� �o���H�j��
Function Ant_data_Function(in_scode,in_no,type_string)
'�o���H�j��
     iSQL = ""
     iSQL = " Select ant_country,ant_cname1,ant_cname2,ant_ename1,ant_ename2,ant_fcname,ant_lcname,ant_fename,ant_lename from dmp_ant "
     iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
        'response.Write isql 
        'response.End
     rs1.Open iSQL,conn,1,1
     
     if not rs1.EOF then
          for i=1 to rs1.RecordCount 
              Ant_data_local = Ant_data()
              iisql = ""
              iisql = " select b.coun_code,b.coun_cname,b.coun_ename From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code "
              iisql = iisql & " where a.coun_code = '"&trim(rs1("ant_country"))&"'"
            Country_name = empty
            rs2.Open iisql,conn,1,1
            if not rs2.EOF then 
                Country_name =  trim(rs2("coun_code"))&trim(rs2("coun_cname"))
            end if
               rs2.Close
            SELECT CASE type_string
            CASE "IG_E"                       
                Ant_data_local = ReplaceData(Ant_data_local, "#ant_num#","�o���H"&i,"empty")
            CASE "UG_E"
                Ant_data_local = ReplaceData(Ant_data_local, "#ant_num#","�s���Ч@�H"&i,"empty")
            CASE "DG1_E"
                Ant_data_local = ReplaceData(Ant_data_local, "#ant_num#","�]�p�H"&i,"empty")
            END SELECT            
            
            Ant_data_local = ReplaceData(Ant_data_local, "#ant_country#", Country_name,"empty")
           
            Cname_string = trim(rs1("ant_fcname"))&","&trim(rs1("ant_lcname")) 
            if Cname_string = "," then 
               Cname_string =  trim(rs1("ant_cname1"))&trim(rs1("ant_cname2"))
            end if
            Ename_string = trim(rs1("ant_fename"))&","&trim(rs1("ant_lename")) 
            if Ename_string = "," then 
               Ename_string =  trim(rs1("ant_ename1"))&trim(rs1("ant_ename2"))
            end if
          
            Ant_data_local = ReplaceData(Ant_data_local, "#ant_cname#", Cname_string,"empty")
            
            Ant_data_local = ReplaceData(Ant_data_local, "#ant_ename#", ToXmlUnicode(Ename_string),"empty")
            
            
            objStream.WriteText = Ant_data_local
            rs1.MoveNext 
             '�ťզ�
            objStream.WriteText = Space_String
          NEXT
     END IF
     rs1.Close
End function

'20170524 �W�[�q�l���ک��Y�ﶵ
Function Dmp_rectitle_Function(in_scode,in_no,receipt_title)
	'�ӽФH�j��.
	iSQL = ""
	iSQL = "select b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename"
	iSQL = iSQL & " from dmp_apcust a,apcust b "
	iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'  and kind='A' and a.apsqlno=b.apsqlno"
	'response.Write isql 
	'response.End
	rs1.Open iSQL,conn,1,1
	
	
	if not rs1.EOF then
		'for i=1 to rs1.RecordCount 
		'���ꤽ�q
		if Mid(trim(rs1("apclass")),1,1) = "A" Then
			Cname_string = trim(rs1("ap_cname1"))&trim(rs1("ap_cname2")) 
		end if
		'����۵M�H
		if Mid(trim(rs1("apclass")),1,1) = "B" Then
			Cname_string = trim(rs1("ap_fcname"))&trim(rs1("ap_lcname")) 
			if Cname_string = "" then 
				Cname_string =  trim(rs1("ap_cname1"))&trim(rs1("ap_cname2"))
			end if
		end if
			'20161206�~��H/���q�W�[�P�_-�Y����g�ӽФH�m&�W,�h��ܩm�W
		if Mid(trim(rs1("apclass")),1,1) = "C" Then
			Cname_string = trim(rs1("ap_fcname"))&trim(rs1("ap_lcname")) 
			if Cname_string = "" then 
				Cname_string =  trim(rs1("ap_cname1"))&trim(rs1("ap_cname2"))
				else
			end if
		end if
			
		'rs1.MoveNext 
		
		'Next
	end if
	rs1.Close
	
	Title_string=Cname_string
	if receipt_title="A" then'�M�Q�v�H
		Title_string=Cname_string
	elseif receipt_title="C" then'�M�Q�v�H(�Nú�H)
		Title_string=Title_string&"(�Nú�H�G�t�q��ڱM�Q�Ӽ��p�X�ưȩ�)"
	elseif receipt_title="B" then'�ť�
		Title_string=""
	end if
	
	Dmp_rectitle_Function=Title_string
	'Dmp_receipt_title_local = Dmp_receipt_title()
	'Dmp_receipt_title_local = ReplaceData(Dmp_receipt_title_local, "#rectitle_name#", ToUnicode2(Title_string),"empty")
	'objStream.WriteText = Dmp_receipt_title_local
End function 


'�򥻸�ƪ� �ӽФH�j��
Function Dmp_apcust_data_1_function(in_scode,in_no)
   '�򥻸��:�ӽФH�j��.
     iSQL = ""
     iSQL = "select b.ap_zip,b.ap_crep,b.ap_erep,b.ap_eaddr1,b.ap_eaddr2,b.ap_eaddr3,b.ap_eaddr4,b.ap_addr1,b.ap_addr2"
     ISQL = ISQL & ",b.apcust_no,b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename"
     iSQL = iSQL & " from dmp_apcust a,apcust b "
     iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'  and kind='A' and a.apsqlno=b.apsqlno"
     'response.Write isql  &"<BR>"
     'response.End
     rs1.Open iSQL,conn,1,1
     
     if not rs1.EOF then
          for i=1 to rs1.RecordCount 
            Dmp_apcust_data_1_local = Dmp_apcust_data_1()
            iisql = ""
            iisql = " select b.coun_code,b.coun_cname,b.coun_ename From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code "
            iisql = iisql & " where a.coun_code = '"&trim(rs1("ap_country"))&"'"
            'REsponse.Write iisql
            'response.End
            Country_name = empty
            rs2.Open iisql,conn,1,1
            if not rs2.EOF then 
                Country_name =  trim(rs2("coun_code"))&trim(rs2("coun_cname"))
            end if
            rs2.Close
            
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#apply_num#", i,"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_country#", Country_name,"empty")
                       
             '���ꤽ�q
            'response.Write trim(rs1("apclass")) &"<BR>"
            if Mid(trim(rs1("apclass")),1,1) = "A" Then
                if trim(rs1("apclass"))="AD" then
                    String_apclass = "�Ӹ��渹�u�t"
                else
                    String_apclass = "�k�H���q�����Ǯ�"
                end if
                Title_cname = "����W��"
                Title_ename = "�^��W��"
                Cname_string = trim(rs1("ap_cname1"))&trim(rs1("ap_cname2")) 
                Ename_string = trim(rs1("ap_ename1"))&trim(rs1("ap_ename2")) 
            end if
            '����۵M�H
            if Mid(trim(rs1("apclass")),1,1) = "B" Then
                String_apclass = "�۵M�H"
                Title_cname = "����m�W"
                Title_ename = "�^��m�W"
                Cname_string = trim(rs1("ap_fcname"))&","&trim(rs1("ap_lcname")) 
                if Cname_string = "," then 
                    Cname_string =  trim(rs1("ap_cname1"))&trim(rs1("ap_cname2"))
                end if
                Ename_string = trim(rs1("ap_fename"))&","&trim(rs1("ap_lename")) 
                if Ename_string = "," then 
                    Ename_string =  trim(rs1("ap_ename1"))&trim(rs1("ap_ename2"))
                end if
            end if
			'20161206�~��H/���q�W�[�P�_-�Y����g�ӽФH�m&�W,�h��ܩm�W
            if Mid(trim(rs1("apclass")),1,1) = "C" Then
                Cname_string = trim(rs1("ap_fcname"))&","&trim(rs1("ap_lcname")) 
                if Cname_string = "," then 
					String_apclass = "�k�H���q�����Ǯ�/�Ӹ��渹�u�t"
					Title_cname = "����W��"
					Title_ename = "�^��W��"
                    Cname_string =  trim(rs1("ap_cname1"))&trim(rs1("ap_cname2"))
				else
					String_apclass = "�۵M�H"
					Title_cname = "����m�W"
					Title_ename = "�^��m�W"
                end if
                Ename_string = trim(rs1("ap_fename"))&","&trim(rs1("ap_lename")) 
                if Ename_string = "," then 
                    Ename_string =  trim(rs1("ap_ename1"))&trim(rs1("ap_ename2"))
                end if
            end if      
                  
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_class#", String_apclass,"empty")
            objStream.WriteText = Dmp_apcust_data_1_local
            
            if trim(rs1("ap_country")) = "T" then
                'response.Write trim(rs1("apcust_no")) &"<BR>"
                Dmp_apcust_data_1_local = ""
                Dmp_apcust_data_1_local = Dmp_apcust_data_1_2
                Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#apcust_no#", trim(rs1("apcust_no")),"empty")
                objStream.WriteText = Dmp_apcust_data_1_local
            end if

            Dmp_apcust_data_1_local = ""
            Dmp_apcust_data_1_local = Dmp_apcust_data_1_3
            
                'response.Write Title_cname &"<BR>"
                'response.Write Title_ename &"<BR>"
                'response.Write Cname_string &"<BR>"
                'response.Write Ename_string &"<BR>"
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_cname1_title#", Title_cname,"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_country#", Country_name,"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_ename1_title#", Title_ename,"empty")                   
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_cname1#", ToUnicode2(Cname_string),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_ename1#", ToXmlUnicode(Ename_string),"empty")
      
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_zip#", trim(rs1("ap_zip")),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_addr1#", ToUnicode2(trim(rs1("ap_addr1"))),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_addr2#", ToUnicode2(trim(rs1("ap_addr2"))),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_eaddr1#", ToXmlUnicode(trim(rs1("ap_eaddr1"))),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_eaddr2#", ToXmlUnicode(trim(rs1("ap_eaddr2"))),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_eaddr3#", ToXmlUnicode(trim(rs1("ap_eaddr3"))),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_eaddr4#", ToXmlUnicode(trim(rs1("ap_eaddr4"))),"empty")
            
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_crep#", trim(rs1("ap_crep")),"empty")
            Dmp_apcust_data_1_local = ReplaceData(Dmp_apcust_data_1_local, "#ap_erep#", ToXmlUnicode(trim(rs1("ap_erep"))),"empty")
      
            objStream.WriteText = Dmp_apcust_data_1_local
            
           
            rs1.MoveNext 
             '�ťզ�
            objStream.WriteText = Space_String
         
          Next
     end if
     rs1.Close
END function

'�򥻸�ƪ� �N�z�H1
Function Agt_data_1_Function(in_scode,in_no)
    ' 201609�P����T�{�A�אּ�Ӧ���줧�N�z�H
    iSQL = ""
    iSQL = " Select b.agt_fax,b.agt_tel,b.agt_addr,b.agt_zip,b.agt_id1,b.agt_id2,b.agt_idno1,b.agt_idno2,b.agt_name1,b.agt_name2 from dmp a "
    iSQL = iSQL & " inner join vdmpall c on a.dmp_sqlno = c.dmp_sqlno "
    iSQL = iSQL & " inner join agt b on c.nagt_no = b.agt_no"
    iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
    'response.Write isql 
    'response.End
    rs1.Open iSQL,conn,1,1  
    '�N�z�H1           
    Agt_data_3_local = Agt_data_3()
    if len(trim(rs1("agt_idno1")))<5 then
        agt_idno1 = string(5-len(trim(rs1("agt_idno1"))),"0") & trim(rs1("agt_idno1"))
    else
        agt_idno1 = trim(rs1("agt_idno1"))
    end if
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_idno1#", agt_idno1,"empty")
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_id1#", trim(rs1("agt_id1")),"empty")
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_name1#", Mid(trim(rs1("agt_name1")),1,1)&","&mid(trim(rs1("agt_name1")),2,2),"empty")
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_zip#", trim(rs1("agt_zip")),"empty")
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_addr#", trim(rs1("agt_addr")),"empty")
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_tel#", trim(rs1("agt_tel")),"empty")
    Agt_data_3_local   = ReplaceData(Agt_data_3_local,   "#agt_fax#", trim(rs1("agt_fax")),"empty")

    objStream.WriteText = Agt_data_3_local

    '�ťզ�
    objStream.WriteText = Space_String
    rs1.close
END Function

'�򥻸�ƪ� �N�z�H2
Function Agt_data_2_Function(in_scode,in_no)
    ' 201609�P����T�{�A�אּ�Ӧ���줧�N�z�H
    iSQL = ""
    iSQL = " Select b.agt_fax,b.agt_tel,b.agt_addr,b.agt_zip,b.agt_id1,b.agt_id2,b.agt_idno1,b.agt_idno2,b.agt_name1,b.agt_name2 from dmp a "
    iSQL = iSQL & " inner join vdmpall c on a.dmp_sqlno = c.dmp_sqlno "
    iSQL = iSQL & " inner join agt b on c.nagt_no = b.agt_no"
    iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
    'response.Write isql 
    'response.End
    rs1.Open iSQL,conn,1,1  
    '�N�z�H2
    Agt_data_4_local = Agt_data_4()  
    if len(trim(rs1("agt_idno2")))<5 then
        agt_idno2 = string(5-len(trim(rs1("agt_idno2"))),"0") & trim(rs1("agt_idno2"))
    else
        agt_idno2 = trim(rs1("agt_idno2"))
    end if
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_idno2#", agt_idno2,"empty")
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_id2#", trim(rs1("agt_id2")),"empty")
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_name2#", Mid(trim(rs1("agt_name2")),1,1)&","&mid(trim(rs1("agt_name2")),2,2),"empty") 
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_zip#", trim(rs1("agt_zip")),"empty")
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_addr#", trim(rs1("agt_addr")),"empty")
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_tel#", trim(rs1("agt_tel")),"empty")
    Agt_data_4_local   = ReplaceData(Agt_data_4_local,   "#agt_fax#", trim(rs1("agt_fax")),"empty")
    objStream.WriteText = Agt_data_4_local     
    '�ťզ�
    objStream.WriteText = Space_String
    rs1.Close
END Function

'
Function Ant_data_1_Function(in_scode,in_no,type_string)
  '�򥻸�ƪ� : �o���H�j��
     iSQL = ""
     iSQL = " Select ant_id,ant_country,ant_cname1,ant_cname2,ant_ename1,ant_ename2,ant_fcname,ant_lcname,ant_fename,ant_lename from dmp_ant "
     iSQL = iSQL & " where in_scode='" & in_scode & "' and in_no='" & in_no & "'"
        'response.Write isql 
        'response.End
     rs1.Open iSQL,conn,1,1
     
     if not rs1.EOF then
          for i=1 to rs1.RecordCount 
              Ant_data_1_local = Ant_data_1()
              Ant_data_1_2_local = Ant_data_1_2()
              iisql = ""
              iisql = " select b.coun_code,b.coun_cname,b.coun_ename From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code "
              iisql = iisql & " where a.coun_code = '"&trim(rs1("ant_country"))&"'"
            
            'REsponse.Write iisql
            'response.End
            
            rs2.Open iisql,conn,1,1
            if not rs2.EOF then 
                Country_name = empty
                Country_name =  trim(rs2("coun_code"))&trim(rs2("coun_cname"))
            end if
               rs2.Close
            SELECT CASE type_string
            CASE "IG_E"                       
                Ant_data_1_local = ReplaceData(Ant_data_1_local, "#ant_num#","�o���H"&i,"empty")
            CASE "UG_E"
                Ant_data_1_local = ReplaceData(Ant_data_1_local, "#ant_num#","�s���Ч@�H"&i,"empty")
            CASE "DG1_E"
                Ant_data_1_local = ReplaceData(Ant_data_1_local, "#ant_num#","�]�p�H"&i,"empty")
            END SELECT
                                    
            Ant_data_1_local = ReplaceData(Ant_data_1_local, "#ant_country#", Country_name,"empty")
           
            Cname_string = trim(rs1("ant_fcname"))&","&trim(rs1("ant_lcname")) 
            if Cname_string = "," then 
               Cname_string =  trim(rs1("ant_cname1"))&trim(rs1("ant_cname2"))
            end if
            Ename_string = trim(rs1("ant_fename"))&","&trim(rs1("ant_lename")) 
            if Ename_string = "," then 
               Ename_string =  trim(rs1("ant_ename1"))&trim(rs1("ant_ename2"))
            end if
          
            Ant_data_1_2_local = ReplaceData(Ant_data_1_2_local, "#ant_cname#", Cname_string,"empty")
            Ant_data_1_2_local = ReplaceData(Ant_data_1_2_local, "#ant_ename#", ToXmlUnicode(Ename_string),"empty")
            objStream.WriteText = Ant_data_1_local
            
            if trim(rs1("ant_country")) = "T" then
                Ant_data_1_1_local = Ant_data_1_1()
                Ant_data_1_1_local = ReplaceData(Ant_data_1_1_local, "#ant_id#", trim(rs1("ant_id")),"empty")    
                objStream.WriteText = Ant_data_1_1_local
            end if
                      
            objStream.WriteText = Ant_data_1_2_local
            rs1.MoveNext 
             '�ťզ�
            objStream.WriteText = Space_String
          NEXT
     END IF
     rs1.Close
END FUNCTION
 %>