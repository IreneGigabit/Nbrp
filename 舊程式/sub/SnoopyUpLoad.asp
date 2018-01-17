<%
Dim SnoopyUpload_SourceData

Class SnoopyUpload
	Dim objForm, objFile, Version
	Dim CharsetEncoding

	Public Function Form(strForm)
		strForm = LCase(strForm)
		If Not objForm.Exists(strForm) Then
			Form = ""
		Else
			Form = objForm(strForm)
		End If
	End Function

	Public Function File(strFile)
		strFile = LCase(strFile)
		If Not objFile.exists(strFile) Then
			Set File = New FileInfo
		Else
			Set File = objFile(strFile)
		End If
	End Function

	Public Sub UploadInit(charset)
		Dim RequestData, sStart, Crlf, sInfo, iInfoStart, iInfoEnd, tStream, iStart, theFile
		Dim iFileSize, sFilePath, sFileType, sFormValue, sFileName
		Dim iFindStart, iFindEnd
		Dim iFormStart, iFormEnd, sFormName

		Version = "Upload Width Progress Bar Version 1.0"
		Set objForm = Server.CreateObject("Scripting.Dictionary")
		Set objFile = Server.CreateObject("Scripting.Dictionary")
		If Request.TotalBytes < 1 Then Exit Sub
		Set tStream = Server.CreateObject("ADODB.Stream")
		Set SnoopyUpload_SourceData = Server.CreateObject("ADODB.Stream")
		SnoopyUpload_SourceData.Type = 1
		SnoopyUpload_SourceData.Mode =3
		SnoopyUpload_SourceData.Open

		Dim TotalBytes
		Dim ChunkReadSize
		Dim DataPart, PartSize

		TotalBytes = Request.TotalBytes     '總大小
		ChunkReadSize = 64 * 1024    ' 分塊大小64K
		BytesRead = 0
		CharsetEncoding = charset
		If CharsetEncoding = "" Then CharsetEncoding = "utf-8"
		'循環分塊讀取
		Do While BytesRead < TotalBytes
			'分塊讀取
			PartSize = ChunkReadSize
			If PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			SnoopyUpload_SourceData.Write DataPart
		Loop

		'SnoopyUpload_SourceData.Write  Request.BinaryRead(Request.TotalBytes)
		SnoopyUpload_SourceData.Position = 0
		RequestData = SnoopyUpload_SourceData.Read

		iFormStart = 1
		iFormEnd = LenB(RequestData)
		Crlf = chrB(13) & chrB(10)
		sStart = MidB(RequestData,1, InStrB(iFormStart, RequestData, Crlf) - 1)
		iStart = LenB(sStart)
		iFormStart = iFormStart + iStart + 1

		While (iFormStart + 10) < iFormEnd
			iInfoEnd = InStrB(iFormStart, RequestData, Crlf & Crlf) + 3
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			SnoopyUpload_SourceData.Position = iFormStart
			SnoopyUpload_SourceData.CopyTo tStream, iInfoEnd - iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.Charset = CharsetEncoding
			sInfo = tStream.ReadText
			tStream.Close
			'取得表單項目名稱
			iFormStart = InStrB(iInfoEnd, RequestData, sStart)
			iFindStart = InStr(22, sInfo, "name=""", 1) + 6
			iFindEnd = InStr(iFindStart, sInfo, """", 1)
			sFormName = LCase(Mid (sinfo, iFindStart, iFindEnd - iFindStart))
			'如果是文件
			If InStr (45, sInfo, "filename=""", 1) > 0 Then
				Set theFile = new FileInfo
				'取得文件名
				iFindStart = InStr(iFindEnd, sInfo, "filename=""", 1) + 10
				iFindEnd = InStr(iFindStart, sInfo, """", 1)
				sFileName = Mid(sinfo, iFindStart, iFindEnd - iFindStart)
				theFile.FileName = getFileName(sFileName)
				theFile.FileExt = getFileExt(sFileName)
				theFile.FilePath = getFilePath(sFileName)
				'取得文件類型
				iFindStart = InStr(iFindEnd,sInfo, "Content-Type: ", 1) + 14
				iFindEnd = InStr(iFindStart, sInfo, vbCr)
				theFile.FileType =Mid(sinfo, iFindStart, iFindEnd - iFindStart)
				theFile.FileStart = iInfoEnd
				theFile.FileSize = iFormStart - iInfoEnd - 3
				theFile.FormName = sFormName
				If Not objFile.Exists(sFormName) Then
					objFile.add sFormName, theFile
				End If
			Else
				'如果是表單項目
				tStream.Type = 1
				tStream.Mode =3
				tStream.Open
				SnoopyUpload_SourceData.Position = iInfoEnd 
				SnoopyUpload_SourceData.CopyTo tStream, iFormStart - iInfoEnd - 3
				tStream.Position = 0
				tStream.Type = 2
				tStream.Charset = CharsetEncoding
				sFormValue = tStream.ReadText 
				tStream.Close
				If objForm.Exists(sFormName) Then
					objForm(sFormName) = objForm(sFormName) & ", " & sFormValue          
				Else
					objForm.Add sFormName, sFormValue
				End If
			End If
			iFormStart=iFormStart + iStart + 1
		Wend
		RequestData = ""
		Set tStream = Nothing
	End Sub

	Private Sub Class_Initialize 

	End Sub

	Private Sub Class_Terminate  
		If Request.TotalBytes > 0 Then
			objForm.RemoveAll
			objFile.RemoveAll
			Set objForm = Nothing
			Set objFile = Nothing
			SnoopyUpload_SourceData.Close
			Set SnoopyUpload_SourceData = Nothing
		End If
	End Sub

	Private Function getFilePath(FullPath)
		If FullPath <> "" Then
			getFilePath = Left(FullPath, InStrRev(FullPath, ""))
		Else
			getFilePath = ""
		End If
	End Function

	Private Function getFileExt(FullPath)
		Dim n
		Dim sExt

		If FullPath <> "" Then
			n = InStrRev(FullPath, ".")
			If n > 0 Then
				sExt = Mid(FullPath, n)
			Else
				sExt = ""
			End If
			getFileExt = LCase(sExt)
		Else
			getFileExt = ""
		End If
	End Function

	Private Function getFileName(FullPath)
		If FullPath <> "" Then
			getFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
		Else
			getFileName = ""
		End If
	End Function
End Class

Class FileInfo
	Dim FormName, FileExt, FileName, FilePath, FileSize, FileType, FileStart

	Private Sub Class_Initialize 
		FileName = ""
		FileExt = ""
		FilePath = ""
		FileSize = 0
		FileStart= 0
		FormName = ""
		FileType = ""
	End Sub

	Public Function SaveAs(FullPath)
		Dim dr, ErrorChar, i

		SaveAs = True
		If Trim(fullpath) = "" Or FileStart = 0 Or fileName = "" Or Right(fullpath, 1)= "/" Then Exit Function
		Set dr = CreateObject("ADODB.Stream")
		dr.Mode = 3
		dr.Type =1
		dr.Open
		SnoopyUpload_SourceData.position = FileStart
		SnoopyUpload_SourceData.copyto dr, FileSize
		Session("ExtMsg") = FullPath
		dr.SaveToFile FullPath, 2
		dr.Close
		Set dr = Nothing 
		SaveAs = False
	End Function

	Public Function BinaryToHex()
		Dim sOut, bAy, i, OneByte

		SnoopyUpload_SourceData.position = FileStart
		bAy = SnoopyUpload_SourceData.Read(FileSize)
		sOut = "0x"
		For i = 1 To LenB(bAy)
			OneByte = Right("0" & Hex(AscB(MidB(bAy, i, 1))), 2)
			sOut = sOut & OneByte
		Next
		BinaryToHex = sOut
	End Function
End Class

%>
