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
On Error Resume Next '***���i�����D�`�Y���A�D���|�ݯdprocess
Dim appWord,myDoc,objFSO

response.Write("$('#chkmsg').html('');"&vbcrlf)
'response.Write("$('#chkmsg').append('"&request("catch_path")&"<BR>');"&vbcrlf)
'response.Write("$('#chkmsg').append('"&replace(Server.MapPath(request("catch_path")),"\","\\")&"<BR>');"&vbcrlf)

FileName=Server.MapPath(request("catch_path"))
Set objFSO = CreateObject( "Scripting.FileSystemObject" )
If not objFSO.FileExists( FileName ) Then
	response.Write("$('#chkmsg').html('<Font align=left color=""red"" size=3>�䤣�컡����Word��("&replace(FileName,"\","\\")&")!!</font><BR>');"&vbcrlf)
	Response.end()
end if

'response.Write("$('#chkmsg').append('����컡����Word��<BR>');"&vbcrlf)

Set appWord = CreateObject("Word.Application")
'appWord.Documents.Open FileName
appWord.Documents.Open FileName , , True
'Set myDoc = appWord.ActiveDocument
appWord.Visible = True

response.Write("var errFlag=false;"&vbcrlf)

'����K�n(�q�l�ӽЮ榡)
summary=Get_eSummary()
if summary="" then
	'����K�n(�ȥ��ӽЮ榡)
	summary=Get_pSummary()
end if
summary=replace(replace(summary,"\","\\"), "'","\'")
Response.Write("document.getElementById('summary_text').innerHTML = '"&summary&"';"&vbcrlf)
if summary="" then
	response.Write("errFlag=true;"&vbcrlf)
	response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>�䤣��K�n!!</font><BR>');"&vbcrlf)
end if

'����M�Q�ӽнd��(�q�l�ӽЮ榡)
range=Get_ERange()
if range="" then
	'����M�Q�ӽнd��(�ȥ��ӽЮ榡)
	range=Get_PRange()
end if
range=replace(replace(range,"\","\\"), "'","\'")
Response.Write("document.getElementById('range_text').innerHTML = '"&range&"';"&vbcrlf)
if range="" then
	response.Write("errFlag=true;"&vbcrlf)
	response.Write("$('#chkmsg').append('<Font align=left color=""red"" size=3>�䤣��M�Q�ӽнd��!!</font><BR>');"&vbcrlf)
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
'response.write("	$('#chkmsg').html('<Font align=left color=""darkblue"" size=3>�^�������A�нT�{���e!!</font><BR>');"&vbcrlf)
response.write("	$('#summary_text').focus();"&vbcrlf)
response.write("	alert('�^�������A�нT�{���e!!');"&vbcrlf)
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

'����K�n(�q�l�ӽЮ榡)
Function Get_eSummary()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "�i*�K�n�j"
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('�����11');")
			.Selection.MoveRight wdCharacter, 1
			With .Selection.Find
				.Text = "�i����j"
				.Forward = True
				.MatchWholeWord = True
			End With
		
			If .Selection.Find.Execute Then
				'response.write("alert('�����22');")
				.Selection.MoveRight wdCharacter, 1
				i=0
				Do While i < 100 '����L���j��
					i = i + 1
					.Selection.MoveDown wdParagraph, 1 'ctrl+�U
					.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+�U
					
					strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
					If InStr(strTemp, "�i�^��j") > 0 Or InStr(strTemp, "�i���w�N��ϡj") > 0 Then
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

'����K�n(�ȥ��ӽЮ榡)
Function Get_pSummary()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "����*�K�n�G"
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('�����11');")
			.Selection.MoveRight wdCharacter, 1
			i=0
			Do While i < 100 '����L���j��
				i = i + 1
				.Selection.MoveDown wdParagraph, 1 'ctrl+�U
				.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+�U
				
				strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
				If InStr(strTemp, "�^��") > 0 And InStr(strTemp, "�K�n�G") > 0 Then
					Exit Do
				Else
					get_value = get_value & strTemp
				End If
			Loop
		end if
	End With
	
	Get_pSummary  = Unicode2Htm(get_value)
End function

'����M�Q�ӽнd��(�q�l�ӽЮ榡)
Function Get_ERange()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "�i*�ӽбM�Q�d��j"
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('�����11');")
			i=0
			Do While i < 100 '����L���j��
				i = i + 1
				.Selection.MoveDown wdParagraph, 1 'ctrl+�U
				.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+�U
				
				strTemp = Trim(Replace(Replace(.Selection.Text, Chr(13), "") ,Chr(9), ""))
				If .Selection.Paragraphs(1).Range.ListFormat.ListString <> "�i��1���j" And (InStr(.Selection.Paragraphs(1).Range.ListFormat.ListString, "�i") > 0 Or strTemp = "") Then
					Exit Do
				Else
					get_value = get_value & strTemp
				End If
			Loop
		end if
	End With
	
	Get_ERange  = Unicode2Htm(get_value)
End function

'����M�Q�ӽнd��(�ȥ��ӽЮ榡)
Function Get_PRange()
	get_value  = ""
	appWord.Selection.HomeKey wdStory
	appWord.Selection.Find.ClearFormatting
	With appWord
		With .Selection.Find
			.Text = "�ӽбM�Q�d��G"
			.Forward = True
			.MatchWholeWord = True
			.MatchWildcards = True
		End With
		
		If .Selection.Find.Execute Then
			'response.write("alert('�����11');")
			i=0
			Do While i < 100 '����L���j��
				i = i + 1
				.Selection.MoveDown wdParagraph, 1 'ctrl+�U
				.Selection.MoveDown wdParagraph, 1, wdExtend 'ctrl+shift+�U
				
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
