<%
%>
<html>
<head>
<title><%=cont%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<script language="VBscript">
function AttachFile()
	attachfilename=AttachForm.theFile.value
	if len(attachfilename) = 0 then
		alert "�п�J�n�W�Ǫ��ɮצW�١A�Ψϥ��s���ӿ���ɮסC"
		exit function
	End IF
	AttachForm.hidFile.value = AttachForm.theFile.value
	AttachForm.button1.disabled = true
	document.AttachForm.submit()
End function

</script>
</head>

<%
response.buffer=true

gdept = "P"
select case request.querystring("type")
  case "photo"
      session("type")="photo"
	  session("seq")=Request("seq")
      session("cust_area")=left(Request.QueryString("cust_area"),1)& gdept
      session("draw_file")=Request("draw_file")
      '93/03/08_jessica�ק�
      session("form_name")=Request("form_name")
      '93/03/08_end
      cont="���ɤW��"
  case "doc"
      session("type")="doc"
	  session("folder_name")=Request("folder_name")	'�ɮץؿ��Ҧb,�|�bcust_area���U�إ�
	  session("prefix_name")=Request("prefix_name")	'�ɮת��e�m�W��:�Ҧp: filename=abc.jpg, if prefix="123" , then filename=123_abc.jpg,�Ω�Ϲj�P�ؿ����U,���P���ɮצW��
      session("cust_area")=left(Request.QueryString("cust_area"),1) & gdept
      session("draw_file")=Request("draw_file")	'���ɮ׸��|(�s�b��server�W��D:\data\document)
      session("form_name")=Request("form_name")
        'Response.Write session("form_name") & "<BR>"
      session("size_name")=Request("size_name")
      session("file_name")=Request("file_name")
        'Response.Write session("file_name") & "<BR>"
      session("source_name")=Request("source_name")
      session("filename_flag")=Request("filename_flag")
      session("btnname")=Request("btnname")
      session("nfilename")=Request("nfilename")
      session("doc_in_date")=Request("in_date")
      session("doc_in_scode")=Request("in_scode")
      session("branch_name")=Request("branch_name")
      session("docbranch")=Request("branch")
      session("prgid")=Request("prgid")
      session("tablename")=Request("tablename")
      session("seq")=Request("seq")
      session("seq1")=Request("seq1")
      session("pic_sqlno")=Request("pic_sqlno")
	'	Response.Write "prgid="& session("prgid") & "<BR>"
	'	Response.Write session("prefix_name") & "<BR>"
 	'	Response.Write session("nfilename") & "<BR>"
     'Response.End 
      cont="�ɮפW��"
  case "Ext_photo"
      session("type")="photo"
      session("seq")=Request("seq")
      session("cust_area")=left(Request.QueryString("cust_area"),1)&"PE"
      session("draw_file")=Request("draw_file")
      '93/03/08_jessica�ק�
      session("form_name")=Request("form_name")
      '93/03/08_end
      cont="���ɤW��"
   case "apcust_file","custdb_file","db_file","custresp_file","brdb_file"
      '2012/5/2�W�[�Acustdb_file=��ʱb�Ȩ� db_file=�дڳ� custresp_file=�Ȥ��ʦ^�����
      '2015/11/16 apcust_file �����ѡB�e����
	  '2016/11/22�W�[brdb_file=�^��invoice
      session("type")=request.querystring("type")
	  session("folder_name")=Request("folder_name")	'�ɮץؿ��Ҧb,�|�bcust_area���U�إ�
	  session("prefix_name")=Request("prefix_name")	'�ɮת��e�m�W��:�Ҧp: filename=abc.jpg, if prefix="123" , then filename=123_abc.jpg,�Ω�Ϲj�P�ؿ����U,���P���ɮצW��
      session("cust_area")=left(Request.QueryString("cust_area"),1) & gdept
      session("draw_file")=Request("draw_file")	'���ɮ׸��|(�s�b��server�W��D:\data\document)
      session("form_name")=Request("form_name")
      session("size_name")=Request("size_name")
      session("file_name")=Request("file_name")
      session("source_name")=Request("source_name")
      session("filename_flag")=Request("filename_flag")
      session("btnname")=trim(Request("btnname"))
      session("nfilename")=Request("nfilename")
      session("doc_in_date")=Request("in_date")
      session("doc_in_scode")=Request("in_scode")
      session("doc_in_scodenm")=Request("in_scodenm")
      session("db_file_flag")=Request("db_file_flag")
      session("prgid_name")=Request("prgid_name")
      session("prgid")=Request("prgid")
      session("attach_flag_name")=request("attach_flag_name")
      session("ar_no")=request("ar_no")	'for �^��invoice
      session("qs_dept")=request("qs_dept")	'for �^��invoice
      session("draw_name")=request("draw_name")	'for �^��invoice
      cont="�ɮפW��"    
      if request.querystring("type")="custdb_file" then
         cont="��ʱb�Ȩ��ɮפW��"    
      elseif request.querystring("type")="db_file" then
         cont="�дڳ��ɮפW��"  
      elseif request.querystring("type")="custresp_file" then
         cont="��ʱb�Ȥ�^���ɮפW��"
      elseif request.querystring("type")="apcust_file" then
        session("db_file_flag") = ""
      elseif request.querystring("type")="brdb_file" then
        cont="�^��Invoice�ɮפW��"
        'session("db_file_flag") = ""
      end if
  case else
      session("type")=""
      response.write "<html><head><title>RE?!1ORE!?DAo3?�go�g!!C</title></head><body bgcolor=#ffffff><br><br><p><center>RE?!1ORE!?DAo3?�go�g!!C"
      response.write "<form><input type=button value=Ao3?�go�g! onclick=""javascript:parent.close()""></form></center></body></html>"
      Response.End
end select

%>
<body bgcolor="#FFFFFF">
<p align="center"><big><font face="�з���" color="#004000"><strong><big><big><%=cont%></big></big></strong></font></big></p>

<center>
  <form id="AttachForm" name="AttachForm" action="upload_winact_file_new.asp" method="Post" enctype="multipart/form-data">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr><td>
        <div align="left">
          �@�W���ɮר쥻������:
          <br>
          �@<input type="file" id="theFile" name="theFile" size="25">
          �@<input type="hidden" id="hidFile" name="hidFile" size="25">
          �@<input type="hidden" id="hidoverwrite" name="hidoverwrite" size="25">
          �@<input type="hidden" id="nfilename" name="nfilename" size="<%=Request("nfilename")%>">
			<input type="hidden" id="tablename" name="tablename" size="<%=Request("tablename")%>">
			<input type="hidden" id="testshow" name="testshow" size="<%=Request("folder_name")%>">
          <br>&nbsp;
<span style="display:none">          
<font size="2" color="red">
<input type="checkbox" id="chkoverwrite" name="chkoverwrite">�л\�w�s�b���ɮ�<br></font></span>
          <br>
          <table width="95%" border="0">
            <tr> 
              <td>
<!--<font color="red"><strong>[�Ъ`�N]</strong>�W���ɮ��������o�ۧ@�v�H���ѭ��P�N�A���o���I�ǥL�H�ۧ@�v���欰!</font><br><br>-->
<div align="center"><center>

<table border="0" width="100%">
    <tr>
        <td align="left" colspan=2  style="font:14px;color:#009900">�ϥΤ覡�G</td>
    </tr>
    <tr><td width="4%" align="right" valign="top">��</td>
        <td width="96%" style="font:14px;color:black">
        ���W���ɮצܥ����A���I��W�褧�y�s���z���s��|�X�{�@�ӡy����ɮסz�p�����A�M��п�ܱz�q�������W�Ǥ��ɮסC
        </td>
    </tr>
    <tr><td width="4%" align="right" valign="top">��</td>
        <td width="96%" style="font:14px;color:red">
        �Y�ݤW��Word�ɡA�Щ��x�s�ɮ׮ɡA�ɮ�������ܡuWord 97-2003���(*.doc)�v�A�H�K����˵��L�k�}�ҡC
        </td>
    </tr>
<!--  <tr>
    <td width="9%" align="right" valign="top"><font size="2" color="black">��</td>
    <td width="91%"><font size="2" color="red">�ɮצW���i�H������r�C</td>
  </tr>-->
</table>
  </center></div>
</td></tr>
</table>
        </div>
</td></tr>
<tr><td align="center">
<input type="button" value="�W��" onclick="AttachFile()" id="button1" name="button1" class="cbutton">
<input type="button" value="��������" onclick="javascript:parent.close()" id="button2" name="button2" class="cbutton">
</td></tr>
  </table>
</form>
</center></body>
</html>
