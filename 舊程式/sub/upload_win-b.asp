<%response.buffer=true
  formcode=request("formcode")  
  form_data=request("data")
  form_title=request("title")
  form_id=request("form_id")
  form_name=request("form_name")
  formmode=mid(formcode,1,1)
  formcopy=mid(formcode,2,1)  
  datatype=mid(formcode,3,1)  
  form=request("form")
  
%>
<html>
<head>
<title><%=form_title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
function AttachFile()
{
  var attachfilename = document.AttachForm.theFile.value;
  if (attachfilename.length == 0) { 
    alert('�п�J�n�W�Ǫ��ɮצW�١A�Ψϥ��s���ӿ���ɮסC');
    return false;
  }
  document.AttachForm.button1.disabled =true;
  document.AttachForm.button2.disabled =true;        
  document.AttachForm.submit();
  return true;
}
</script>
</head>
<body bgcolor="#FFFFFF">
<p align="center"><big><font face="�з���" color="#004000"><strong><%=cont%></strong></font></big></p>

<center>
  <form name="AttachForm" action="upload_winact.asp" method="POST" enctype="multipart/form-data">  
     <input name=form type=hidden value=<%=form%>>
     <input name=form_id type=hidden value=<%=form_id%>>
     <input name=form_name type=hidden value=<%=form_name%>>  
     <input name=form_title type=hidden value=<%=form_title%>> 
     <input name=form_data type=hidden value=<%=form_data%>> 
   <input name=ctrlnum type=hidden value=<%=ctrlnum%>>
   <input name=formcode type=hidden value=<%=formcode%>>    
   <input name=datatype type=hidden value=<%=datatype%>>   

    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr><td>
        <div align="left">
          �@�W��<font size="2" color="red">���</font>�ɮר쥻������:
          <br>
          �@<input type="file" name="theFile" size="25">
          <br><br>
          <table width="95%" border="0">
            <tr> 
              <td>
<font color="red"><strong>[�Ъ`�N]</strong>�W���ɮ��������o�ۧ@�v�H���ѭ��P�N�A���o���I�ǥL�H�ۧ@�v���欰!</font><br><br>
<font size="2" color="#009900">�ϥΤ覡�G</font><br>
<font size="2" color="black">
�����W���ɮצܥ����A���I��W�褧�y�s���z���s��|�X�{�@�ӡy�ɮ׿���z�p�����A�M��п�ܱz�q�������W�Ǥ��ɮסC</font>
</font>
</td></tr>
</table>
        </div>
</td></tr>
<tr><td align="center">
<input type="button" value="�W��" onclick="AttachFile()" id="button1" name="button1">
<input type="button" value="��������" onclick="javascript:parent.close()" id="button2" name="button2">
</td></tr>
  </table>
</form>
</center></body>
</html>



