<%
Response.CharSet = "BIG5"
Session.CodePage = 950

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 20

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open session("btbrtdb")
Set rs0 = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
SQL=unescape(Request("SearchSql"))
%>
String.prototype.HTMLEncode = function(str) {
	var result = "";
	var str = (arguments.length===1) ? str : this;
	for(var i=0; i<str.length; i++) {
		var chrcode = str.charCodeAt(i);
		result+=(chrcode>128) ? "&#"+chrcode+";" : str.substr(i,1);
	}
	return result;
}
<%
On Error Resume Next '***不可拿掉非常嚴重，主機會殘留process
Dim appWord,myDoc,objFSO

SQL="select * from dmp_attach "
SQL = SQL & "where seq = '"& request("seq") &"' "
SQL = SQL & "and seq1 = '"& request("seq1") &"' "
SQL = SQL & "and step_grade = '"& request("step_grade") &"' "
SQL = SQL & "and attach_flag<>'D' "
SQL = SQL & "and esend_flag='' "
SQL = SQL & "and attach_desc like '%申請書%' "
SQL = SQL & "and source_name like '%.doc' "
rs0.Open SQL,conn,1,1

response.Write("$('#chkmsg').html('');")
if rs0.eof then
	'response.Write "<Font align=left color='red' size=3>找不到申請書Word檔，請先上傳!!</font>" & "<BR>"
	response.Write("$('#chkmsg').html('<Font align=left color=""red"" size=3>找不到申請書Word檔，請先上傳!!〈word檔判斷規則：副檔名為.doc，附件說明含有「申請書」字樣，不可勾□電子送件檔〉</font><BR>');"&vbcrlf)
	Response.end()
elseif rs0.recordcount>1 then
	'response.Write "<Font align=left color='red' size=3>找到多個申請書Word檔，請確認!!</font>" & "<BR>"
	response.Write("$('#chkmsg').html('<Font align=left color=""red"" size=3>找到多個申請書Word檔，請確認!!</font><BR>');"&vbcrlf)
	Response.end()
else
	'FileName=Server.MapPath("..\report-word-xml\reportdata\[延展註冊申請書]-ST22985_1.doc")
	FileName=Server.MapPath(rs0("attach_path"))
	'Response.write(FileName&"<BR>")
	Set objFSO = CreateObject( "Scripting.FileSystemObject" )
	If not objFSO.FileExists( FileName ) Then
		'response.Write "<Font align=left color='red' size=3>找不到申請書Word檔("&FileName&")!!</font>" & "<BR>"
		response.Write("$('#chkmsg').html('<Font align=left color=""red"" size=3>找不到申請書Word檔("&replace(FileName,"\","\\")&")!!</font><BR>');"&vbcrlf)
		Response.end()
	end if
	
	Set appWord = CreateObject("Word.Application")
	'appWord.Documents.Open FileName
	appWord.Documents.Open FileName , , True
	'Set myDoc = appWord.ActiveDocument
	appWord.Visible = True
	
	response.Write("var errFlag=false;"&vbcrlf)
	
	'20170808 增加檢查案件名稱
	title_line=Get_name("【")
	title_line=replace(replace(title_line,"【",""),"】","")
	isql = " select form_name from cust_code where Code_type='word_tit_p' and code_name='"&title_line&"' "
	rs1.open isql,conn,1,1
	if rs1.eof then
		if session("se_scode")="m1583" then
			response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>找不到申請書設定，請聯繫資訊人員!!("&replace(isql,"'","\'")&")</font><BR>');"&vbcrlf)
		else
			response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>找不到申請書設定，請聯繫資訊人員!!</font><BR>');"&vbcrlf)
		end if
	else
		arr_appl=split(rs1("form_name"),"|")'中文專利名稱tag|英文專利名稱tag
		cappl_line=Get_name(arr_appl(0))'中文專利名稱tag
		split_cappl=split(cappl_line,"】")
		eappl_line=Get_name(arr_appl(1))'英文專利名稱tag
		split_eappl=split(eappl_line,"】")
		
		'檢查中文專利名稱
		response.Write("var cappl_name=document.getElementsByName('cappl_name')[0].value"&vbcrlf)
		'response.Write("var cappl_name=$('input[name=cappl_name]').val().html();"&vbcrlf)
		'response.Write("var cappl_name=$('<div/>').text($('input[name=cappl_name]').val()).html();"&vbcrlf)
		response.Write("if (cappl_name.HTMLEncode()!='"&trim(split_cappl(1))&"'.HTMLEncode()){"&vbcrlf)
		response.Write("	errFlag=true;"&vbcrlf)
		if session("se_scode")="m1583" then
			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_cappl(0)&"】申請書案件名稱("&trim(split_cappl(1))&")與案件主檔('+cappl_name+')不符!!</font><BR>');"&vbcrlf)
		else
			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_cappl(0)&"】申請書案件名稱與案件主檔不符!!</font><BR>');"&vbcrlf)
		end if
		response.Write("}"&vbcrlf)
		
		'檢查英文專利名稱(官發確認畫面無此欄位)
		'if UBOUND(split_eappl)>0 then
		'	response.Write("var eappl_name=document.getElementsByName('eappl_name')[0].value"&vbcrlf)
		'	response.Write("if (eappl_name!='"&trim(split_eappl(1))&"'){"&vbcrlf)
		'	response.Write("	errFlag=true;"&vbcrlf)
		'	if session("se_scode")="m1583" then
		'		response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_eappl(0)&"】申請書案件名稱("&trim(split_eappl(1))&")與案件主檔('+eappl_name+')不符!!</font><BR>');"&vbcrlf)
		'	else
		'		response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_eappl(0)&"】申請書案件名稱與案件主檔不符!!</font><BR>');"&vbcrlf)
		'	end if
		'	response.Write("}"&vbcrlf)
		'end if
		
		'20170815改抓官發確認畫面
		'isql = " select cappl_name,eappl_name from dmp a where seq = '"& request("seq") &"' and seq1 = '"& request("seq1") &"' "
		'rs3.open isql,conn,1,1
		'IF rs3.EOF then
		'	response.Write("	errFlag=true;"&vbcrlf)
		'	if session("se_scode")="m1583" then
		'		response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_cappl(0)&"】找不到案件主檔!!(EOF)("&replace(isql,"'","\'")&")</font><BR>');"&vbcrlf)
		'	else
		'		response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_cappl(0)&"】找不到案件主檔!!</font><BR>');"&vbcrlf)
		'	end if
		'else
		'	'檢查中文專利名稱
		'	if trim(rs3("cappl_name"))<>trim(split_cappl(1)) then
		'		response.Write("	errFlag=true;"&vbcrlf)
		'		if session("se_scode")="m1583" then
		'			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_cappl(0)&"】申請書案件名稱("&trim(split_cappl(1))&")與案件主檔("&trim(rs3("cappl_name"))&")不符!!</font><BR>');"&vbcrlf)
		'		else
		'			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_cappl(0)&"】申請書案件名稱與案件主檔不符!!</font><BR>');"&vbcrlf)
		'		end if
		'	end if
		'	'檢查英文專利名稱
		'	if UBOUND(split_eappl)>0 then
		'		if trim(rs3("eappl_name"))<>trim(split_eappl(1)) then
		'			response.Write("	errFlag=true;"&vbcrlf)
		'			if session("se_scode")="m1583" then
		'			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_eappl(0)&"】申請書案件名稱("&trim(split_eappl(1))&")與案件主檔("&trim(rs3("eappl_name"))&")不符!!</font><BR>');"&vbcrlf)
		'			else
		'				response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_eappl(0)&"】申請書案件名稱與案件主檔不符!!</font><BR>');"&vbcrlf)
		'			end if
		'		end if
		'	end if
		'end if
		'rs3.close
	end if
	rs1.close
	
	'20170808 增加檢查規費
	fee_line=Get_name("【繳費金額】")
	split_fee=split(fee_line,"】")
	if UBOUND(split_fee)=1 then
		response.Write("var fee=document.getElementsByName('fees')[0].value"&vbcrlf)
		response.Write("if (fee!='"&trim(split_fee(1))&"'){"&vbcrlf)
		response.Write("	errFlag=true;"&vbcrlf)
		if session("se_scode")="m1583" then
			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>【繳費金額】官發應繳規費('+fee+')與申請書填寫金額("&trim(split_fee(1))&")不符!!</font><BR>');"&vbcrlf)
		else
			response.Write("	$('#chkmsg').append('<Font align=left color=""red"" size=3>【繳費金額】官發應繳規費與申請書填寫金額不符!!</font><BR>');"&vbcrlf)
		end if
		response.Write("}"&vbcrlf)
	end if

	'20170126 原用tagList定義的tag名檢查,改用word【附送書件】區塊,查dmp_attach是否有上傳
	arrBlock=Get_AttachBlock()
	split_attach_Block = split(arrBlock,"#")
	FOR i=0 to UBOUND(split_attach_Block) 
		if split_attach_Block(i)<>"" then
			split_line=split(replace(split_attach_Block(i),"　",""),"】")
			if UBOUND(split_line)=1 then
				'response.write("alert('"&split_line(0)&"】="&split_line(1)&"');"&vbcrlf)
				isql = " select * from dmp_attach a "
				isql = isql & "where seq = '"& request("seq") &"' "
				isql = isql & " and seq1 = '"& request("seq1") &"' "
				isql = isql & " and step_grade = '"& request("step_grade") &"' "
				isql = isql & " and source_name='"& trim(split_line(1)) &"' "
				isql = isql & " and esend_flag='Y' "
				isql = isql & " and attach_flag<>'D' "
				if session("se_scode")="m1583" then
				'response.write(isql&"<BR>")
				end if
				'response.write(isql&"<BR>"&vbcrlf)
				rs2.open isql,conn,1,1
				
				IF rs2.EOF then
					'errFlag=true
					'response.Write "<Font align=left color='red' size=3>"&split_line(0)&"】<b>"&split_line(1)&"</b> 抓取對應附件有錯誤，請檢查附送書件之檔案是否已經上傳 !!</font>" & "<BR>"
					response.Write("errFlag=true;"&vbcrlf)
					response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>"&split_line(0)&"】<b>"&split_line(1)&"</b> 抓取對應附件有錯誤，請檢查附送書件之檔案是否已經上傳 !!</font><BR>');"&vbcrlf)
				End IF
				rs2.Close
			end if
		end if
	NEXT
	
	if Err.number > 0 Then
		'response.write "<Font align=left color='red' size=3>Eeception - " & ERR.number & ERR.description & "</font>" & "<BR>"
		response.Write("errFlag=true;"&vbcrlf)
		response.Write("$('#chkmsg').html('<Font align=left color=""red"" size=3>Eeception - " & ERR.number & ERR.description & "!!</font><BR>');"&vbcrlf)
	end if
	Set rs0 = Nothing
	Set rs1 = Nothing
	Set rs2 = Nothing
	Set rs3 = Nothing
	'Set WordDoc = Nothing 
	appWord.Quit
	set appWord = Nothing
end if

response.write("if (!errFlag){"&vbcrlf)
response.write("	$('#chkmsg').html('<Font align=left color=""darkblue"" size=3>檢查完成，請執行確認!!</font><BR>');"&vbcrlf)
response.write("	$('#button0').attr('disabled', true);"&vbcrlf)
response.write("}"&vbcrlf)
%>

<%
Const wdCell=12
Const wdCharacter=1
Const wdCharacterFormatting=13
Const wdColumn=9
Const wdItem=16
Const wdLine=5
Const wdParagraph=4
Const wdParagraphFormatting=14
Const wdRow=10
Const wdScreen=7
Const wdSection=8
Const wdSentence=3
Const wdStory=6
Const wdTable=15
Const wdWindow=11
Const wdWord=2
Const wdExtend=1
Const wdMove=0

Function Unicode2Htm(s)
	Dim ret,i,c,a,w
	If IsNull(s) or Trim(s)="" then
		Unicode2Htm = ""
		Exit Function
	End If  
	 
	ret = ""
	for i=1 to Len(s)
		c = Mid(s,i,1)
		a = Asc(c)
		w = Ascw(c)
		If w<0 then
			w = 65536 + w
		End If
		If a=63 and w<>63 then
			ret = ret & "&#" & w & ";"   
		ElseIf w>127 and w<256 then
			ret = ret & "&#" & w & ";"
		Else
			ret = ret & c
		End If
	next
	Unicode2Htm = ret
End Function

Function Get_name(pTag_name)
	get_value  = ""
	appWord.Selection.HomeKey 6
	With appWord.Selection
		'.ClearFormatting
		.Find.Text = pTag_name
		.Find.Forward = True
		.Find.MatchWholeWord = True  
		'.Execute
	
		If .Find.Execute Then
			appWord.Selection.HomeKey 5 
			'appWord.Selection.EndKey 5, 1
			appWord.Selection.MoveDown 4, 1, 1'5,1
			appWord.Selection.Copy
			get_value = trim(replace(appWord.Selection.text,chr(13),""))'整行複製會帶最後的換行符號
			get_value = replace(get_value,"　","")'全形空白
			get_value = replace(get_value,chr(9),"")'tab
			if session("se_scode")="m1583" then
				'response.write("alert('"&Unicode2Htm(get_value)&"');")
			end if
		end if
	End With

	Get_name  = Unicode2Htm(get_value)
End function

'查word【附送書件】區塊,找到具結為止
Function Get_AttachBlock()
	attach_block = ""
	
	appWord.Selection.HomeKey 6
	With appWord.Selection
		'.ClearFormatting
		.Find.Text = "【附送書件】"
		.Find.Forward = True
		.Find.MatchWholeWord = True  
		'.Execute
	End With
	
	i=0
	If appWord.Selection.Find.Execute Then
		Do While i < 100'防止無限迴圈
			i=i+1
			'response.Write i & "<BR>"
			
			appWord.Selection.MoveDown wdParagraph, 1 'ctrl+↓
			appWord.Selection.MoveDown wdParagraph, 1, 1 'ctrl+shift+↓
			appWord.Selection.Copy
			strTemp = replace(appWord.Selection.text,chr(13),"")'整行複製會帶最後的換行符號
			strTemp = replace(strTemp,"　","")'全形空白
			strTemp = replace(strTemp,chr(9),"")'tab
			strTemp = replace(strTemp,chr(12),"")'換頁
			strTemp = trim(strTemp)
			
			if session("se_scode")="m1583" then
				response.write("$('body').append( '"&strTemp&"<BR>' );")
			end if
			
			if instr(strTemp,"【檔案具結】")>0 or strTemp="【本申請書所檢送之PDF檔或影像檔與原本或正本相同】" or strTemp="【本申請書所填寫之資料係為真實】" then
				exit do
			elseif instr(strTemp,"【其他】")>0 or instr(strTemp,"【文件描述】")>0 or strTemp="" or strTemp="【附送書件】" then
				'continue
			else
				strTemp=replace(strTemp,"【文件檔名】","【其他】")
				'response.write("alert('"&strTemp&"');")
				attach_block= attach_block & "#" &strTemp
			end if
		Loop
	end if
	Get_AttachBlock  = attach_block
End Function

%>
