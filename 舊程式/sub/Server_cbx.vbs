<%
Function ShowSelect(pconn,pSQL,pType)
'pType:true-->no_name(代號_名稱), false-->name(名稱)
	'On Error Resume Next
	response.write "<option value='' style='color:blue' selected>請選擇</option>"
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
		else
			response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
		end if
		response.write chr(10)
		tRSa.MoveNext
	wend
	set tRSa = nothing
'	set pconn = nothing
End Function
Function ShowSelect1(pconn,pSQL,pType,pcho)
'pType:true-->no_name(代號_名稱), false-->name(名稱)
	'On Error Resume Next
	if pcho="Y" then
		response.write "<option value='' style='color:blue' selected>請選擇</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
		else
			response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
		end if
		response.write chr(10)
		tRSa.MoveNext
	wend
	set tRSa = nothing
'	set pconn = nothing
End Function

Function ShowSelect2(pconn,pSQL,pType,pcho)
'pType:true-->no_name(代號_名稱), false-->name(名稱)  retrun string
	'On Error Resume Next
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
	ShowSelect2=innerhtml
	'Response.write ShowSelect2
End Function
'組html
Function ShowSelect5(pconn,pSQL,pType,pcho,pvalue)
'pType:true-->no_name(代號_名稱), false-->name(名稱)  retrun string
	'On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>請選擇</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	
	set tRSa = pConn.execute(pSQL)
		
	'if err.number <> 0 then
	'response.write	pSQL
	'response.end	
	'end if
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			end if				
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if				
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing	
	ShowSelect5=innerhtml
	'Response.write ShowSelect5
End Function
'組html
Function ShowSelect7(pconn,pSQL,pType,pcho,pvalue)
'pType:true-->no_name(代號_名稱), false-->name(名稱)  retrun string
	'On Error Resume Next
	if pcho="Y" then
		innerhtml=innerhtml & "<option value='' style='color:blue' selected>請選擇</option>"
	end if
	set tRsa = Server.CreateObject("ADODB.Recordset")
	
	set tRSa = pConn.execute(pSQL)
		
	'if err.number <> 0 then
	'response.write	pSQL
	'response.end	
	'end if
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' value1='" & Trim(tRSa(1).value) & "' selected>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' value1='" & Trim(tRSa(1).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			end if				
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' value1='" & Trim(tRSa(1).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "' value1='" & Trim(tRSa(1).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if				
		end if
		innerhtml=innerhtml 
		tRSa.MoveNext
	wend
	set tRSa = nothing	
	ShowSelect7 = innerhtml
	'Response.write ShowSelect5
End Function

'取得名稱
function getname(pconn,psql)
	getname = ""
	set tRSa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(psql)
	if not tRSa.eof then getname = tRSa(1)
	set tRSa = nothing
end function
'組html
'Function ShowSelect2(pconn,pSQL,pType,pcho)
'pType:true-->no_name(代號_名稱), false-->name(名稱)  retrun string
'	'On Error Resume Next
'	if pcho="Y" then
'		innerhtml=innerhtml & "<option value='' style='color:blue' selected>請選擇</option>"
'	end if
'	set tRsa = Server.CreateObject("ADODB.Recordset")
'	set tRSa = pConn.execute(pSQL)
'	while not tRSa.eof
'		if pType then
'			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
'		else
'			innerhtml=innerhtml & "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
'		end if
'		innerhtml=innerhtml 
'		tRSa.MoveNext
'	wend
'	set tRSa = nothing
'	ShowSelect2=innerhtml
'End Function

Function ShowSelect3(pconn,pSQL,pType,pvalue)
'pType:true-->no_name(代號_名稱), false-->name(名稱)
    'On Error Resume Next
	response.write "<option value='' style='color:blue' selected>請選擇</option>"
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				response.write "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			end if
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				response.write "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if
		end if
		response.write chr(10)
		tRSa.MoveNext
	wend
	set tRSa = nothing
'	set pconn = nothing
End Function

Function ShowSelect4(pconn,pSQL,pType,pvalue,poption)
'pType:true-->no_name(代號_名稱), false-->name(名稱)
    'On Error Resume Next
	response.write "<option value='' style='color:blue' selected>" & poption & "</option>"
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				response.write "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			end if
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				response.write "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if
		end if
		response.write chr(10)
		tRSa.MoveNext
	wend
	set tRSa = nothing
'	set pconn = nothing
End Function
Function ShowSelect6(pconn,pSQL,pType,pvalue)
'pType:true-->no_name(代號_名稱), false-->name(名稱)
    'On Error Resume Next
	response.write "<option value='' style='color:blue' selected>全部</option>"
	set tRsa = Server.CreateObject("ADODB.Recordset")
	set tRSa = pConn.execute(pSQL)
	while not tRSa.eof
		if pType then
			if trim(pvalue)=Trim(tRSa(0).value) then
				response.write "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			else
				response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(0).value) & "_" & Trim(tRSa(1).value) & "</option>"
			end if
		else
			if trim(pvalue)=Trim(tRSa(0).value) then
				response.write "<option value='" & Trim(tRSa(0).value) & "' selected>" & Trim(tRSa(1).value) & "</option>"
			else
				response.write "<option value='" & Trim(tRSa(0).value) & "'>" & Trim(tRSa(1).value) & "</option>"
			end if
		end if
		response.write chr(10)
		tRSa.MoveNext
	wend
	set tRSa = nothing
'	set pconn = nothing
End Function

Function getSetting(pDept,pType)
	select case pType
		case "1"   
			'期限稽催列印規則設定 管制日期為系統日 ~ 系統日 + 2 天後到期的顯示紅色
			'管制日期 <= 2 日內顯示紅色
			select case ucase(pDept)
				case "T"
					getSetting = "d;2;red"
				case "P"
					getSetting = "d;2;red"
			end select
	end select
End Function
%>