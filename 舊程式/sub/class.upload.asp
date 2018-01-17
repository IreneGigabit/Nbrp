<%
Class Upload 
  '公用存取變數
	Dim path        '上傳路徑
	Dim maxSize     '最大容量限制
  '--------------------------------------------------
	'內部存取屬性,方法
  Private upobj
	Private data
	Private binaryData  '上傳的二進位資料
  '--------------------------------------------------
  '開始時執行的工作
  Private Sub Class_Initialize()
	  data = Request.TotalBytes
		binaryData = Request.BinaryRead(data)
		Set upobj   = Server.CreateObject("basp21")
		path = ".\tmp\"
		maxSize = 0  '沒有限制 
  End Sub

  '結束時執行的工作
  Private Sub Class_Terminate()
    'Something to do
  End Sub

  '取得上傳檔案路徑
	Function get_path(file)
    get_path = upobj.FormFileName(binaryData, file) 
	End Function

	'將檔名擷取出來
	Function get_file(file)
     get_file =right(file,len(file)-InstrRev(file,"\"))
	End Function
	'將檔名2擷取出來
	Function get_file1(file)
     get_file1 =upobj.FormFileName(binaryData, file) 
	End Function
	'取得不是 file 欄位的值 
	Function get_fd(fd)
    get_fd = upobj.Form(binaryData, fd) 
	End Function
	'取得檔案大小
	Function get_FileSize(file)
		get_FileSize = upobj.FormFileSize(binarydata, file) 
	End Function
	'將檔案儲存
	Function SaveFile(file)
		Dim saveFileName
		If Right(path, 1)<>"\" then path=path & "\"
		saveFileName = get_file(get_path(file))
		'回傳上傳位元數
		
		SaveFile = upobj.FormSaveAs(binaryData, file, Server.MapPath(path & saveFileName)) 
  End Function
End Class

Function echo(v)
  Response.Write v
End Function
%>