<%
'�s�W Log �ɡA�A�Ω� log table ���� ud_flag�Bud_date�Bud_scode�Bprgid �o������
'ptable�Gex:step_imp �n�s�W�� step_imp_log �h�ǤJ  step_imp
'pkey_filed�Gkey �����W�١A�p���h�����ХΡF�j�}
'pkey_value�G�P pkey_field �ۤ��t�X�A�p���h�����ХΡF�j�}
function insert_log_table(pconn,pud_flag,pprgid,ptable,pkey_field,pkey_value)
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
	'response.write tisql
	'response.end
		
	
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
	tisql = tisql & " select '" & pud_flag & "',getdate(),'" &session("scode") & "',"
	tisql = tisql & "'" &pprgid& "'," & tfield_str
	tisql = tisql & " from " & ptable
	tisql = tisql & " where 1 = 1 "
	tisql = tisql & wsql
	'response.write tisql & "<br>"
	'response.end	
	On Error Resume Next
	pconn.execute tisql
	If err.number<>0 Then Call errorLoggin("�J"& ptable &"log��", tisql, pprgid)
    On Error Goto 0	
	
	set tRS = nothing
end function

' �N���~SQL�g�Jlog��
Sub errorLoggin (mStr, sqlStr, pgID)
	Dim ecnn
	Dim eSQL

	Set ecnn = Server.CreateObject("ADODB.connection")
	ecnn.ConnectionString = session("btbrtdb")
	ecnn.Open

	eSQL = "INSERT INTO [brp_error_log] ([log_scode], [log_date], [prgid], [MsgStr], [SQLstr]) VALUES ("
	eSQL = eSQL & "'"& Session("scode") & "'"
	eSQL = eSQL & "," & "GETDATE()"
	eSQL = eSQL & ",'" & pgID & "'"
	eSQL = eSQL & ",'" & mStr & "'"
	eSQL = eSQL & ",'" & SQLstr & "'"
	eSQL = eSQL & ")"
	'Session("EXtMsg") = eSQL
	ecnn.Execute eSQL

	ecnn.Close
	Set ecnn = Nothing
End Sub
%>
