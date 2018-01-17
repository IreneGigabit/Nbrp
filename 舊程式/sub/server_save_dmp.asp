
<% 
' 優先權資料
Function update_dmp_prior(dmp_sqlno, seq, seq1)
    Set rsf = Server.CreateObject("ADODB.RecordSet") 

    Dim dmp_prior_cnt
    
    For i=1 To Request.Form("priornum")
		If Trim(Request.Form("pri"& i &"_prior_sqlno"))<>Empty then
			If Trim(Request.Form("pri"& i &"_del_flag"))="Y" then
			' 刪除
				usql = "DELETE FROM dmp_prior"
				usql = usql & " OUTPUT 'D', GETDATE(), " & chkcharnull(Session("scode")) &"," & chkcharnull(prgid) & ", DELETED.* INTO dmp_prior_log"
				usql = usql & " WHERE prior_sqlno = '"& Request.Form("pri"& i &"_prior_sqlno") &"'"
				
				If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
				conn.execute(usql)
				If err.number<>0 Then geterrmsg()
			Else
			' 修改				
				usql = "UPDATE dmp_prior SET"                
                usql = usql & " prior_yn = "& chkcharnull(Request.Form("pri"& i &"_prior_yn"))
                usql = usql & ", apply_yn = "& chkcharnull(Request.Form("pri"& i &"_apply_yn"))
                usql = usql & ", prior_no = "& chkcharnull(Request.Form("pri"& i &"_prior_no"))
                usql = usql & ", prior_country = "& chkcharnull(Request.Form("pri"& i &"_prior_country"))
                usql = usql & ", prior_date = "& chkdatenull(Request.Form("pri"& i &"_prior_date"))
                usql = usql & ", prior_seq = "& chkcharnull(Request.Form("pri"& i &"_prior_seq"))
                usql = usql & ", prior_seq1 = "& chkcharnull(Request.Form("pri"& i &"_prior_seq1"))
                usql = usql & ", mprior_access = "& chkcharnull(Request.Form("pri"& i &"_mprior_access"))
                usql = usql & ", prior_case1 = "& chkcharnull(Request.Form("pri"& i &"_prior_case1"))
                usql = usql & ", prior_change_date = "& chkdatenull(Request.Form("pri"& i &"_prior_change_date"))
                usql = usql & ", prior_change_no = "& chkcharnull(Request.Form("pri"& i &"_prior_change_no"))

				usql = usql & ", tran_date = GETDATE()" 
				usql = usql & ", tran_scode = "& chkcharnull(Session("scode"))
                usql = usql & " OUTPUT 'U', GETDATE(), " & chkcharnull(Session("scode")) & "," & chkcharnull(prgid) & ", DELETED.* INTO dmp_prior_log"				
				usql = usql & " WHERE prior_sqlno = '"& Request.Form("pri"& i &"_prior_sqlno") &"'"				
						
				If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
				conn.execute(usql)
				If err.number<>0 Then geterrmsg()	
			End If		
		Else
			If Trim(Request.Form("pri"& i &"_del_flag"))="Y" Then
			
			Else
			' 新增
				If Trim(Request.Form("pri"& i &"_prior_date"))<>Empty And Trim(Request.Form("pri"& i &"_prior_country"))<>Empty Then
				' 若有申請日、國別，則將該筆資料入檔
					usql = "INSERT INTO dmp_prior ("
					usql = usql & " dmp_sqlno, seq, seq1"
					usql = usql & ", prior_yn, apply_yn, prior_no, prior_country, prior_date, prior_seq, prior_seq1, mprior_access, prior_case1, prior_change_date, prior_change_no" 
					usql = usql & ", tran_date, tran_scode"
					usql = usql & ") VALUES ("
					usql = usql & " "& chkcharnull(dmp_sqlno)
					usql = usql & ", "& chkcharnull(seq)
					usql = usql & ", "& chkcharnull(seq1)
					
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_yn"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_apply_yn"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_no"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_country"))
                    usql = usql & ", "& chkdatenull(Request.Form("pri"& i &"_prior_date"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_seq"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_seq1"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_mprior_access"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_case1"))
                    usql = usql & ", "& chkdatenull(Request.Form("pri"& i &"_prior_change_date"))
                    usql = usql & ", "& chkcharnull(Request.Form("pri"& i &"_prior_change_no"))
                    
					usql = usql & ", GETDATE()"	
					usql = usql & ", "& chkcharnull(Session("scode"))
					usql = usql & ")"
							
					If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
					conn.execute(usql)
					If err.number<>0 Then geterrmsg()	
				End If		
			End If
		End If
    Next

    For i=1 To Request.Form("priornum")
        If Trim(Request.Form("pri"& i &"_del_flag"))<>"Y" Then
            If Request.Form("pri"& i &"_prior_yn")="Y" And Request.Form("pri"& i &"_prior_country")="T" And Trim(Request.Form("pri"& i &"_prior_seq"))<>Empty And Trim(Request.Form("pri"& i &"_prior_seq1"))<>Empty Then
                Call insert_log_table(conn, "U", prgid, "dmp", "seq;seq1", Request.Form("pri"& i &"_prior_seq") &";"& Trim(Request.Form("pri"& i &"_prior_seq1")))
                
                usql = "UPDATE dmp SET"
                usql = usql & " fprior_flag = 'Y'"
                'usql = usql & " OUTPUT 'U', GETDATE(), " & chkcharnull(Session("scode")) & "," & chkcharnull(prgid) & ", DELETED.* INTO dmp_log"
                usql = usql & " WHERE seq = '"& Request.Form("pri"& i &"_prior_seq") &"'"
                usql  =usql & " AND seq1 = '"& Request.Form("pri"& i &"_prior_seq1") &"'"
                
			    If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
			    conn.execute(usql)
			    If err.number<>0 Then geterrmsg()	
            End If
        End If
    Next	
            
    For i=1 To Request.Form("priornum")
        If Request.Form("pri"& i &"_oprior_yn")="Y" And Request.Form("pri"& i &"_oprior_country")="T" And Trim(Request.Form("pri"& i &"_oprior_seq"))<>Empty And Trim(Request.Form("pri"& i &"_oprior_seq1"))<>Empty Then
            isql = "SELECT COUNT(*) AS cnt"
            isql = isql & " FROM dmp_prior"
            isql = isql & " WHERE prior_seq = '"& Request.Form("pri"& i &"_oprior_seq") &"'"
            isql = isql & " AND prior_seq1 = '"& Request.Form("pri"& i &"_oprior_seq1") &"'"
            isql = isql & " AND prior_country = 'T'"
            isql = isql & " AND prior_yn = 'Y'"

            If Trim(Request.Form("chkTest"))<>Empty Then Response.Write isql & "<br /><br />"
            rsf.Open isql, conn, 0, 1
            dmp_prior_cnt = rsf("cnt")
            rsf.Close
            
            If CInt(dmp_prior_cnt) = CInt(0) Then
                Call insert_log_table(conn, "U", prgid, "dmp", "seq;seq1", Request.Form("pri"& i &"_oprior_seq") &";"& Trim(Request.Form("pri"& i &"_oprior_seq1")))
                
                usql = "UPDATE dmp SET"
                usql = usql & " fprior_flag = 'N'"
                'usql = usql & " OUTPUT 'U', GETDATE(), " & chkcharnull(Session("scode")) & "," & chkcharnull(prgid) & ", DELETED.* INTO dmp_log"
                usql = usql & " WHERE seq = '"& Request.Form("pri"& i &"_oprior_seq") &"'"
                usql  =usql & " AND seq1 = '"& Request.Form("pri"& i &"_oprior_seq1") &"'"
                
			    If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
			    conn.execute(usql)
			    If err.number<>0 Then geterrmsg()
            End If
        End If
    Next
    
    ' 將優先權日寫回案件主檔
    If seq<>Empty Then
        Call insert_log_table(conn, "U", prgid, "dmp", "seq;seq1", seq &";"& seq1)
        
        usql = "UPDATE dmp SET"
        usql = usql & " prior_date = (SELECT MIN(prior_date) FROM dmp_prior WHERE seq = '"& seq &"' AND seq1 = '"& seq1 &"' AND prior_yn = 'Y')"
        usql = usql & " WHERE seq = '"& seq &"'"
        usql = usql & " AND seq1 = '"& seq1 &"'"
        
	    If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
	    conn.execute(usql)
	    If err.number<>0 Then geterrmsg()        
    Else
        Call insert_log_table(conn, "U", prgid, "dmp", "dmp_sqlno", dmp_sqlno)
        
        usql = "UPDATE dmp SET"
        usql = usql & " prior_date = (SELECT MIN(prior_date) FROM dmp_prior WHERE dmp_sqlno = '"& dmp_sqlno &"' AND prior_yn = 'Y')"
        usql = usql & " WHERE dmp_sqlno = '"& dmp_sqlno &"'"
        
	    If Trim(Request.Form("chkTest"))<>Empty Then Response.Write "<font color=brown>sub\server_save_dmp.asp→</font>"&usql & "<hr />"
	    conn.execute(usql)
	    If err.number<>0 Then geterrmsg()     
    End If

    Set rsf = Nothing
End Function

%>