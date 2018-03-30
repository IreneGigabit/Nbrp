<%
Response.CharSet = "BIG5"
Session.CodePage = 950

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 20
%>

<!--#INCLUDE FILE="../sub/Server_conn_unicode.vbs" -->

<%
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

response.Write("$('#chkmsg').html('');"&vbcrlf)
'response.Write("$('#chkmsg').append('"&request("catch_path")&"<BR>');"&vbcrlf)
'response.Write("$('#chkmsg').append('"&replace(Server.MapPath(request("catch_path")),"\","\\")&"<BR>');"&vbcrlf)

FileName=Server.MapPath(request("catch_path"))
Set objFSO = CreateObject( "Scripting.FileSystemObject" )
If not objFSO.FileExists( FileName ) Then
	response.Write("$('#chkmsg').html('<Font align=left color=""red"" size=3>找不到說明書Word檔("&replace(FileName,"\","\\")&")!!</font><BR>');"&vbcrlf)
	Response.end()
end if

'response.Write("$('#chkmsg').append('有找到說明書Word檔<BR>');"&vbcrlf)

Set appWord = CreateObject("Word.Application")
'appWord.Documents.Open FileName
appWord.Documents.Open FileName , , True
'Set myDoc = appWord.ActiveDocument
appWord.Visible = True

response.Write("var errFlag=false;"&vbcrlf)

'抓取摘要(電子申請格式)
summary=Get_eSummary()
if summary="" then
	'抓取摘要(紙本申請格式)
	summary=Get_pSummary()
end if
summary=replace(replace(summary,"\","\\"), "'","\'")
Response.Write("document.getElementById('summary_text').innerHTML = '"&summary&"';"&vbcrlf)
if summary="" then
	response.Write("errFlag=true;"&vbcrlf)
	response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>找不到摘要!!</font><BR>');"&vbcrlf)
end if

'抓取專利申請範圍(電子申請格式)
range=Get_ERange()
if range="" then
	'抓取專利申請範圍(紙本申請格式)
	range=Get_PRange()
end if
range=replace(replace(range,"\","\\"), "'","\'")
Response.Write("document.getElementById('range_text').innerHTML = '"&range&"';"&vbcrlf)
if range="" then
	response.Write("errFlag=true;"&vbcrlf)
	response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>找不到專利申請範圍!!</font><BR>');"&vbcrlf)
end if


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

response.write("if (!errFlag){"&vbcrlf)
'response.write("	$('#chkmsg').html('<Font align=left color=""darkblue"" size=3>擷取完成，請確認內容!!</font><BR>');"&vbcrlf)
response.write("	$('#summary_text').focus();"&vbcrlf)
response.write("	alert('擷取完成，請確認內容!!');"&vbcrlf)
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

'抓取摘要(電子申請格式)
Function Get_eSummary()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "【*摘要】"
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('有找到11');")
			.Selection.MoveRight wdCharacter, 1
			With .Selection.Find
				.Text = "【中文】"
				.Forward = True
				.MatchWholeWord = True
			End With
		
			If .Selection.Find.Execute Then
				'response.write("alert('有找到22');")
				.Selection.MoveRight wdCharacter, 1
				i=0
				Do While i < 100 '防止無限迴圈
					i = i + 1
					.Selection.MoveDown wdParagraph, 1 'ctrl+下
					.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+下
					
					strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
					If InStr(strTemp, "【英文】") > 0 Or InStr(strTemp, "【指定代表圖】") > 0 Then
						Exit Do
					Else
						get_value = get_value & strTemp
					End If
				Loop
			end if
		end if
	End With
	
	Get_eSummary  = Unicode2Htm(get_value)
End function

'抓取摘要(紙本申請格式)
Function Get_pSummary()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "中文*摘要："
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('有找到11');")
			.Selection.MoveRight wdCharacter, 1
			i=0
			Do While i < 100 '防止無限迴圈
				i = i + 1
				.Selection.MoveDown wdParagraph, 1 'ctrl+下
				.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+下
				
				strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
				If InStr(strTemp, "英文") > 0 And InStr(strTemp, "摘要：") > 0 Then
					Exit Do
				Else
					get_value = get_value & strTemp
				End If
			Loop
		end if
	End With
	
	Get_pSummary  = Unicode2Htm(get_value)
End function

'抓取專利申請範圍(電子申請格式)
Function Get_ERange()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "【*申請專利範圍】"
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('有找到11');")
			i=0
			Do While i < 100 '防止無限迴圈
				i = i + 1
				.Selection.MoveDown wdParagraph, 1 'ctrl+下
				.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+下
				
				strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
				If .Selection.Paragraphs(1).Range.ListFormat.ListString <> "【第1項】" And (InStr(.Selection.Paragraphs(1).Range.ListFormat.ListString, "【") > 0 Or strTemp = "") Then
					Exit Do
				Else
					get_value = get_value & strTemp
				End If
			Loop
		end if
	End With
	
	Get_ERange  = Unicode2Htm(get_value)
End function

'抓取專利申請範圍(紙本申請格式)
Function Get_PRange()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "申請專利範圍："
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('有找到11');")
			i=0
			Do While i < 100 '防止無限迴圈
				i = i + 1
				.Selection.MoveDown wdParagraph, 1 'ctrl+下
				.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+下
				
				strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
				If ((.Selection.Paragraphs(1).Range.ListFormat.ListString <> "1" And .Selection.Paragraphs(1).Range.ListFormat.ListString <> "") Or InStr(strTemp, "2.") > 0 Or strTemp = "") Then
					Exit Do
				Else
					get_value = get_value & strTemp
				End If
			Loop
		end if
	End With
	
	Get_PRange  = Unicode2Htm(get_value)
End function

%>
