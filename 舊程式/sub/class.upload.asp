<%
Class Upload 
  '���Φs���ܼ�
	Dim path        '�W�Ǹ��|
	Dim maxSize     '�̤j�e�q����
  '--------------------------------------------------
	'�����s���ݩ�,��k
  Private upobj
	Private data
	Private binaryData  '�W�Ǫ��G�i����
  '--------------------------------------------------
  '�}�l�ɰ��檺�u�@
  Private Sub Class_Initialize()
	  data = Request.TotalBytes
		binaryData = Request.BinaryRead(data)
		Set upobj   = Server.CreateObject("basp21")
		path = ".\tmp\"
		maxSize = 0  '�S������ 
  End Sub

  '�����ɰ��檺�u�@
  Private Sub Class_Terminate()
    'Something to do
  End Sub

  '���o�W���ɮ׸��|
	Function get_path(file)
    get_path = upobj.FormFileName(binaryData, file) 
	End Function

	'�N�ɦW�^���X��
	Function get_file(file)
     get_file =right(file,len(file)-InstrRev(file,"\"))
	End Function
	'�N�ɦW2�^���X��
	Function get_file1(file)
     get_file1 =upobj.FormFileName(binaryData, file) 
	End Function
	'���o���O file ��쪺�� 
	Function get_fd(fd)
    get_fd = upobj.Form(binaryData, fd) 
	End Function
	'���o�ɮפj�p
	Function get_FileSize(file)
		get_FileSize = upobj.FormFileSize(binarydata, file) 
	End Function
	'�N�ɮ��x�s
	Function SaveFile(file)
		Dim saveFileName
		If Right(path, 1)<>"\" then path=path & "\"
		saveFileName = get_file(get_path(file))
		'�^�ǤW�Ǧ줸��
		
		SaveFile = upobj.FormSaveAs(binaryData, file, Server.MapPath(path & saveFileName)) 
  End Function
End Class

Function echo(v)
  Response.Write v
End Function
%>