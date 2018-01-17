

<%

      '埃
      'DBDelete(DB_DSN,DB_Table,FD_SQL)    
      Function DBDelete(ByVal temp_DSN,ByVal DB_Table,ByVal FD_SQL)

         dim SQL_temp

         SQL_temp = "DELETE FROM " & DB_Table & " Where " & FD_SQL
         'response.write sql_temp
         'response.end     
      
         DBDelete = DB_Execute(temp_DSN,SQL_temp)

       End Function       



        '穝糤
       'DBInsert(DB_DSN,DB_Table,FD_Name,FD_Value)
       Function DBInsert(ByVal temp_DSN,ByVal DB_Table,ByVal DB_Field,ByVal DB_Value)

          dim SQL_temp
          dim Low_Index
          dim High_Index
          dim Counter
  
          Low_Index = lbound(DB_Field)
          High_Index = ubound(DB_Field) - 1

          SQL_temp = "INSERT INTO " & DB_Table & " ("
          for Counter = Low_Index to High_Index - 1
              SQL_temp = SQL_temp & DB_Field(Counter) & ","
          next
          SQL_temp = SQL_temp & DB_Field(High_Index) & ") VALUES (" 
          for Counter = Low_Index to High_Index - 1
              SQL_temp = SQL_temp & ISNumber(DB_Value(Counter)) & ","
          next
          SQL_temp = SQL_temp & ISNumber(DB_Value(High_Index)) & ")" 
          'DBInsert = SQL_temp
    	'Response.Write SQL_temp
	'Response.End 
 
         DBInsert = DB_Execute(temp_DSN,SQL_temp)
     End Function

     Function DBUpdateM(ByVal temp_DSN,ByVal DB_Table,ByVal DB_Field,ByVal DB_Value,ByVal FD_SQL)
       Set tempRs = Server.CreateObject("ADODB.RecordSet") 

       dim SQL_temp
       dim Low_Index
       dim High_Index
       dim Counter

       Low_Index = lbound(DB_Field)
       High_Index = ubound(DB_Field) - 1

       SQL_temp = "UPDATE " & DB_Table & " SET "

       for Counter = Low_Index to High_Index - 1
           SQL_temp = SQL_temp & DB_Field(Counter) & "=" & ISNumber(DB_Value(Counter)) & ","
       next

       SQL_temp = SQL_temp & DB_Field(High_Index) & "=" & ISNumber(DB_Value(High_Index)) 

       SQL_temp = SQL_temp & " Where " & FD_SQL 
       'RESPONSE.WRITE SQL_TEMP
       'RESPONSE.END  
       
'==============
       strIP = Request.ServerVariables("REMOTE_ADDR")
       'strMac = GetMACAddress(strIP)
       sqllog=" select * from  "&DB_Table&" where "& FD_SQL

       tempRs.Open sqllog, temp_DSN, 3
       'set tempRs=DB_Query1(temp_DSN,sqllog) 
 
       sqllogu=""
       for ui=0 to tempRs.Fields.Count - 1
           sqllogu=sqllogu&tempRs(ui).name&","   
       next 
       sqllogu=mid(sqllogu,1,len(sqllogu)-1)
       tempRs.Close  
   
       sqllog = " insert into "&DB_Table&"_log(ud_flag,ud_date,ud_scode,ud_mac,ud_sql,"
       sqllog = sqllog & sqllogu &")"
       sqllog = sqllog & "select 'U' ,getdate(),'"& session("se_scode") &"','"& strIP &"','"& "" &"',"
       sqllog = sqllog & sqllogu &" from "&DB_Table&" where "& FD_SQL
       'response.write sqllog
       'response.end           
       CALL DB_Execute(temp_DSN,sqllog)  
'=======================

          ' 磅︽SQL
       'response.write sql_temp
       'response.end  
       DBUpdateM = DB_Execute(temp_DSN,SQL_temp)
       set tempRs= nothing         
     End Function


     'э
     'DBUpdate(DB_DSN,DB_Table,FD_Name,FD_Value,FD_SQL)  

     Function DBUpdate(ByVal temp_DSN,ByVal DB_Table,ByVal DB_Field,ByVal DB_Value,ByVal FD_SQL)
 
       dim SQL_temp
       dim Low_Index
       dim High_Index
       dim Counter
       Set tempRs = Server.CreateObject("ADODB.RecordSet")    
'==============

'=======================

       Low_Index = lbound(DB_Field)
       High_Index = ubound(DB_Field) - 1

       SQL_temp = "UPDATE " & DB_Table & " SET "

       for Counter = Low_Index to High_Index - 1
           SQL_temp = SQL_temp & DB_Field(Counter) & "=" & ISNumber(DB_Value(Counter)) & ","
       next

       SQL_temp = SQL_temp & DB_Field(High_Index) & "=" & ISNumber(DB_Value(High_Index)) 

       SQL_temp = SQL_temp & " Where " & FD_SQL 
          ' 磅︽SQL
       'response.write sql_temp
       'response.end  
       DBUpdate = DB_Execute(temp_DSN,SQL_temp)
       set tempRs= nothing         
     End Function



     Function DBUpdateD(ByVal temp_DSN,ByVal DB_Table,ByVal DB_Field,ByVal DB_Value,ByVal FD_SQL)
 
       dim SQL_temp
       dim Low_Index
       dim High_Index
       dim Counter
       strIP = Request.ServerVariables("REMOTE_ADDR")
       strMac = GetMACAddress(strIP)
       sqllog=" select * from  "&DB_Table&" where "& FD_SQL
       set rslog=DB_Query1(temp_DSN,sqllog)  
       sqllogu=""
       for i=0 to rslog.Fields.Count - 1
           sqllogu=sqllogu&rslog(i).name&","   
       next 
       sqllogu=mid(sqllogu,1,len(sqllogu)-1)

   
       sqllog = " insert into "&DB_Table&"_log(ud_flag,ud_date,ud_scode,ud_mac,"
       sqllog = sqllog & sqllogu &")"
       sqllog = sqllog & "select 'D' ,'" & date() &"','"& session("se_scode") &"','"& strMac &"',"
       sqllog = sqllog & sqllogu &" from "&DB_Table&" where "& FD_SQL

       CALL DB_Execute(temp_DSN,sqllog)   

       Low_Index = lbound(DB_Field)
       High_Index = ubound(DB_Field) - 1

       SQL_temp = "UPDATE " & DB_Table & " SET "

       for Counter = Low_Index to High_Index - 1
           SQL_temp = SQL_temp & DB_Field(Counter) & "=" & ISNumber(DB_Value(Counter)) & ","
       next

       SQL_temp = SQL_temp & DB_Field(High_Index) & "=" & ISNumber(DB_Value(High_Index)) 

       SQL_temp = SQL_temp & " Where " & FD_SQL 
          ' 磅︽SQL
       'response.write sql_temp
       'response.end  
       DBUpdated = DB_Execute(temp_DSN,SQL_temp)
      
     End Function

        '磅︽
	Function DB_Execute(temp_conn,strSQL)
		'Dim DB_Conn
		'On Error Resume Next
	
		'Set DB_Conn = Server.CreateObject("ADODB.Connection")
		
		'DB_Conn.CursorLocation =adUseClient 
		
		'DB_Conn.Open DB_DSN
                'DB_Conn.BeginTrans
                 'if session("scode") ="admin" then
   	         '   response.write strsql
                 'end if 
                'response.end  

		Set DB_Execute = temp_conn.Execute(strSQL)
		If Err.Number <> 0 Then
                        'DB_Conn.RollbackTrans   
			'Response.Write "SQL Error" & "<br>" & vbNewLine
			'Response.Write "Error:" & Err.Description & "<br>" & vbNewLine
			'Response.Write strSQL
			'Response.End
                else
                           
                        'DB_Conn.CommitTrans     
		End If
		
 	End Function






        '琩高 
	Function DB_Query(strSQL)
		Dim Rs
	
		On Error Resume Next
		
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		'Rs.CursorLocation = adUseClient 
		
		Rs.Open strSQL, DB_DSN, 3
		
		If Err.Number <> 0 Then
			Response.Write "SQL Error" & "<br>" & vbNewLine
			Response.Write "Error:" & Err.Description & "<br>" & vbNewLine
			Response.Write strSQL
			Response.End
		End If
		
		Set Rs.ActiveConnection = Nothing
		
		Set DB_Query = Rs
		
	End Function

'琩高 
	Function DB_Query1(temp_conn,strSQL)
		Dim Rs
	
		On Error Resume Next
		
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		'Rs.CursorLocation = adUseClient 
		
		Rs.Open strSQL, temp_conn, 3
		
		If Err.Number <> 0 Then
			Response.Write "SQL Error" & "<br>" & vbNewLine
			Response.Write "Error:" & Err.Description & "<br>" & vbNewLine
			Response.Write strSQL
			Response.End
		End If
		
		Set Rs.ActiveConnection = Nothing
		
		Set DB_Query1 = Rs
		
	End Function
	'干0
	'sStr = 饼干ぇ跑计
	'nLen = 羆
	Function AddZero(sStr, nLen)
		
		Dim nNowLen
		Dim sTemp
		Dim i
		
		nNowLen = Len(Trim(sStr))
		
		sTemp = ""
		If nNowLen < nLen Then
			For i = 1 To (nLen - nNowLen)
				sTemp = sTemp & "0"
			Next
		End If
		
		sTemp = sTemp & sStr
		AddZero = sTemp
		
	End Function
	
	'干フ
	'sStr = 饼干ぇ跑计
	'nLen = 羆
	'sLR = 1:干オ  2:干
	Function AddSpace(sStr, nLen, sLR)
		
		Dim nNowLen
		Dim sTemp
		
		nNowLen = Len(Trim(sStr))
		
		sTemp = ""
		
		If CInt(NowLen) < CInt(nLen) Then
			If CInt(sLR) = 1 Then
				sTemp = Space(nLen-nNowLen) & sStr
			ElseIf CInt(sLR) = 2 Then
				sTemp = sStr & Space(nLen-nNowLen)
			End If
		End If
		
		AddSpace = sTemp
		
	End Function

        '琌计
         Function ISNumber(ByVal temp_Value)
         if trim(temp_Value)<>empty then
                   
                if isdate(temp_Value) then
                   if IsNumeric(temp_Value) then
               	     ISNumber = "'" & temp_Value & "'"
                   else   
                      ISNumber = "'" & gettime1(temp_Value) & "'" 
                   end if 
                else   
       		  ISNumber = "'" & temp_Value & "'"
                end if
                  
	 else
		ISNumber = "null"
	 end if  
        End Function	

        '眔丁
         function gettime()
             gettime=date()&" "&hour(time())&":"&Minute(time())&":"&Second(time()) 
         end function             

         function gettime1(ByVal temp_Value)
             gettime1=DateValue(temp_Value)&" "&hour(temp_Value)&":"&Minute(temp_Value)&":"&Second(temp_Value) 
         end function             


         '肚絏 
         Function FPWCODE(BYVAL temp_table,ByVal temp_file,ByVal temp_Value)
         Set connm = Server.CreateObject("ADODB.Connection")
         connm.Open application("DB_DSN")
         sql="select " & temp_file & " from " & temp_table & temp_Value
         'response.write sql
         'response.end    
         set srs=DB_Query(sql)
           
         IF not sRS.EOF then   
            fpwcode = srs(srs(0).name)
         else
            fpwcode = ""
         end if 
         sRs.Close
         Set sRs = Nothing            
         connm.Close             
         Set connm = Nothing
         End Function

         sub readdata3(table,field,whe)
           
             sql="select " & field & " from " & table & whe
             'response.write sql
             'response.end
             set TEMP_rs=DB_Query(sql)
             while not TEMP_rs.EOF
             %>
              <option value = '<% =TEMP_rs(TEMP_rs(1).name) %>' ><% =TEMP_rs(TEMP_rs(0).name) %></option>
             <%
             TEMP_rs.movenext
              wend
              TEMP_rs.close
             

         end sub




function GetMACAddress(strIP)
Set net = Server.CreateObject("wscript.network")
Set sh = Server.CreateObject("wscript.shell")
sh.run "%comspec% /c nbtstat -A " & strIP & " > c:\" & strIP & ".txt",0,true
Set sh = nothing
Set fso = createobject("scripting.filesystemobject")
Set ts = fso.opentextfile("c:\" & strIP & ".txt")
macaddress = null
Do While Not ts.AtEndOfStream
data = ucase(trim(ts.readline))
if instr(data,"MAC ADDRESS") Then
macaddress = trim(split(data,"=")(1))
Exit Do
End if
loop
ts.close
Set ts = nothing
fso.deletefile "c:\" & strIP & ".txt"
Set fso = nothing
GetMACAddress = macaddress
End function 

%>



