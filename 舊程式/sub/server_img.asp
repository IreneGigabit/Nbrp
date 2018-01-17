<%
' 圖片顯示-起始宣告
'sid:圖檔編碼，sext:副檔名
Function DocPictHead(sid, sext)
'DocPictHead = "<w:p " & _
'"wsp:rsidR=""00AD0DFE"" wsp:rsidRDefault=""00F95955""><w:r><w:pict><v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" " & _
'"o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f""><v:stroke joinstyle=""miter""/><v:formulas><v:f " & _
'"eqn=""if lineDrawn pixelLineWidth 0""/><v:f eqn=""sum @0 1 0""/><v:f eqn=""sum 0 0 @1""/><v:f eqn=""prod @2 1 2""/><v:f eqn=""prod " & _
'"@3 21600 pixelWidth""/><v:f eqn=""prod @3 21600 pixelHeight""/><v:f eqn=""sum @0 0 1""/><v:f eqn=""prod @6 1 2""/><v:f eqn=""prod " & _
'"@7 21600 pixelWidth""/><v:f eqn=""sum @8 21600 0""/><v:f eqn=""prod @7 21600 pixelHeight""/><v:f eqn=""sum @10 21600 0""/></v:formulas>" & _
'"<v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/><o:lock v:ext=""edit"" aspectratio=""t""/></v:shapetype><w:binData " & _
'"w:name=""wordml://0300"& sid &"."& sext &""">"

DocPictHead = "<w:p wsp:rsidR=""00CE591C"" wsp:rsidRPr=""005D6A10"" wsp:rsidRDefault=""00CE591C"" wsp:rsidP=""006C3521""><w:pPr><w:snapToGrid " & _
"w:val=""off""/><w:jc w:val=""right""/><w:rPr><w:rFonts w:ascii=""Verdana"" w:h-ansi=""Verdana""/><wx:font wx:val=""Verdana""/></w:rPr>" & _
"</w:pPr></w:p><w:p wsp:rsidR=""001A4692"" wsp:rsidRPr=""005D6A10"" wsp:rsidRDefault=""00101725"" wsp:rsidP=""006C3521""><w:pPr><w:snapToGrid " & _
"w:val=""off""/><w:jc w:val=""center""/><w:rPr><w:rFonts w:ascii=""Verdana"" w:h-ansi=""Verdana""/><wx:font wx:val=""Verdana""/></w:rPr>" & _
"</w:pPr><w:r wsp:rsidRPr=""00101725""><w:rPr><w:rFonts w:ascii=""Verdana"" w:h-ansi=""Verdana""/><wx:font wx:val=""Verdana""/><w:b/>" & _
"</w:rPr><w:pict><v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" " & _
"filled=""f"" stroked=""f""><v:stroke joinstyle=""miter""/><v:formulas><v:f eqn=""if lineDrawn pixelLineWidth 0""/><v:f eqn=""sum @0 " & _
"1 0""/><v:f eqn=""sum 0 0 @1""/><v:f eqn=""prod @2 1 2""/><v:f eqn=""prod @3 21600 pixelWidth""/><v:f eqn=""prod @3 21600 pixelHeight""/>" & _
"<v:f eqn=""sum @0 0 1""/><v:f eqn=""prod @6 1 2""/><v:f eqn=""prod @7 21600 pixelWidth""/><v:f eqn=""sum @8 21600 0""/><v:f eqn=""prod " & _
"@7 21600 pixelHeight""/><v:f eqn=""sum @10 21600 0""/></v:formulas><v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>" & _
"<o:lock v:ext=""edit"" aspectratio=""t""/></v:shapetype>"

DocPictHead = DocPictHead &"<w:binData w:name=""wordml://0200" & sid &"."& sext &""">"

End Function

' 圖片顯示-將圖片轉成base64 code
Function DocPictBody(sFilename)
	Dim objXMLDoc, objDocElem, objStream, sBase64String

	Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument")
	objXMLDoc.async = False
	objXMLDoc.validateOnParse = False

	'sFilename = "/brp/scandoc/NP/test.jpg"
'	sFilename = "/brp/NP/_/218/21880/test.JPG"  文件上傳區
'   sFilename = "/brp/NPE_PIC/NP/test.jpg"  '製圖
	sFilename = Server.MapPath(sFilename)  '轉成實際路徑
	'response.Write sFilename & "<BR>"	
    'response.End 
    
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Type = 1  '二進位方式傳送
	objStream.Open
	'response.Write sFilename &"<BR>"
	'response.End 
	objStream.LoadFromFile sFilename    '載入檔案

	Set objDocElem = objXMLDoc.createElement("pic64")
	objDocElem.dataType = "bin.base64"
	objDocElem.nodeTypedValue = objStream.Read

	sBase64String = objDocElem.text

	objStream.Close

	Set objStream = Nothing
	Set objDocElem = Nothing
	Set objXMLDoc = Nothing

	DocPictBody = sBase64String
End Function

' 圖片顯示-設定寬/高＆結束宣告
'fn : 實際路徑及檔名
'sid : 0000.jpg 圖檔編碼含副檔名
Function DocPictTail(fn, sid, title, md, pic_size)
    Dim ILIB
    Dim iW, iH, oH
    Dim rate
    Dim stmp

    'Set ILIB = server.createobject("Overpower.ImageLib")
    'ILIB.PictureSize fn, iW, iH
    'Set ILIB = Nothing

    Set ILIB = server.createobject("SnopImg.ImgInfo")
    
    Dim x
    Dim dpm
    Dim dp_rate
    Dim ext   
    
    'response.Write fn &"<BR>"
    'response.End 
    
    x = ILIB.SetFile(fn)
    ext = ILIB.ExtType   ' 取得副檔名(.xxx)
    
    If x > 0 Then
        iW = ILIB.Width
        iH = ILIB.Height 
        
        dpm = ILIB.DPMx                ' 取得寬度的dpm (標準值:2835，2835 dpm = 72 dpi)
        dp_rate =  CDbl(2835) / dpm      ' 取得100%縮放比    
    End If
    
    Set ILIB = Nothing

    'If md = "0" Then
    '	oH = iH * 1.28
    '	' Width: 421.85pt / Height:395pt
    '
    '	If iW > oH Then
    '		rate = CSng(418 / iW)
    '	Else
    '		rate = CSng(327 / iH)
    '	End If
    '	'rate = CSng(418 / iW) ' 僅以寬度為基準
    'Else
    '	oH = iH * 0.76068
    '	' Width: 534pt / Height:702pt
    '
    '	If iW > oH Then
    '		rate = CSng(526 / iW)
    '	Else
    '		rate = CSng(692 / iH)
    '	End If
    '	'rate = CSng(526 / iW) ' 僅以寬度為基準
    'End If

    ' 商標圖檔應當不大於10×10cm(283.2pt * 283.2pt)，不小於5×5cm(141.6pt * 141.6pt)
'    If iW > iH Then
'    ' 以寬度為基準
'        If iW>283.2 Then
'            rate = dp_rate * CSng(283.2 / iW)
'        Else
'            rate = dp_rate
'        End If
'    Else
'    ' 以高度為基準
'        If iH>283.2 Then
'            rate = dp_rate * CSng(283.2 / iH)
'        Else
'            rate = dp_rate
'        End If
'    End If

    'rate = rate * 100 / 133
    iW = CInt(iW * rate)
    iH = CInt(iH * rate)
    if pic_size="A4" then
        iW = 520
        iH = 750
'        if iW>cint(520) then iW = 520
'        if iH>cint(750) then iH = 750
    end if
    'width = 357.75pt
    'height = 336.75pt

    'DocPictTail = "</w:binData><v:shape id=""_x0000_i" & Left(sid, 4) & """ type=""#_x0000_t75"" style=""width:" & CStr(iW) & "pt;height:" & CStr(iH) & "pt"">" & _
	'    "<v:imagedata src=""wordml://0300" & sid & """ o:title=""" & title & """/></v:shape></w:pict></w:r></w:p>"

    DocPictTail = "</w:binData><v:shape id=""_x0000_i" & Left(sid, 4) & """ type=""#_x0000_t75"" style=""width:" & CStr(iW) & "pt;height:" & CStr(iH) & "pt"">" & _
	    "<v:imagedata src=""wordml://0200" & sid & """ o:title=""" & title & """/></v:shape></w:pict></w:r></w:p>"

    ' for test
'    DocPictTail = DocPictTail & "<w:p wsp:rsidR=""00D753E7"" wsp:rsidRDefault=""00D753E7"" wsp:rsidP=""00D753E7"">" & _
'        "<w:pPr><w:rPr><w:lang w:fareast=""ZH-CN""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
'        "<w:lang w:fareast=""ZH-CN""/></w:rPr><w:t>檔名："& CStr(fn) &" </w:t></w:r></w:p>"       
'    DocPictTail = DocPictTail & "<w:p wsp:rsidR=""00D753E7"" wsp:rsidRDefault=""00D753E7"" wsp:rsidP=""00D753E7"">" & _
'        "<w:pPr><w:rPr><w:lang w:fareast=""ZH-CN""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
'        "<w:lang w:fareast=""ZH-CN""/></w:rPr><w:t>寬："& CStr(iW) &" pt</w:t></w:r></w:p>"
'    DocPictTail = DocPictTail & "<w:p wsp:rsidR=""00D753E7"" wsp:rsidRDefault=""00D753E7"" wsp:rsidP=""00D753E7"">" & _
'        "<w:pPr><w:rPr><w:lang w:fareast=""ZH-CN""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
'        "<w:lang w:fareast=""ZH-CN""/></w:rPr><w:t>高："& CStr(iH) &" pt</w:t></w:r></w:p>"
'    DocPictTail = DocPictTail & "<w:p wsp:rsidR=""00D753E7"" wsp:rsidRDefault=""00D753E7"" wsp:rsidP=""00D753E7"">" & _
'        "<w:pPr><w:rPr><w:lang w:fareast=""ZH-CN""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
'        "<w:lang w:fareast=""ZH-CN""/></w:rPr><w:t>dp縮放比："& CStr(dp_rate) &" </w:t></w:r></w:p>"      
'    DocPictTail = DocPictTail & "<w:p wsp:rsidR=""00D753E7"" wsp:rsidRDefault=""00D753E7"" wsp:rsidP=""00D753E7"">" & _
'        "<w:pPr><w:rPr><w:lang w:fareast=""ZH-CN""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
'        "<w:lang w:fareast=""ZH-CN""/></w:rPr><w:t>實際縮放比："& CStr(rate/dp_rate) &" </w:t></w:r></w:p>"   

End Function
%>
