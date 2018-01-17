<html>
<head>
<title><%=cont%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<script language="VBscript">
function AttachFile()
	attachfilename=AttachForm.theFile.value
	if len(attachfilename) = 0 then
		alert "請輸入要上傳的檔案名稱，或使用瀏覽來選擇檔案。"
		exit function
	End IF
	AttachForm.hidFile.value = AttachForm.theFile.value
	AttachForm.button1.disabled = true
	document.AttachForm.submit()
End function

</script>
</head>

<%response.buffer=true

gdept = "P"
select case request.querystring("type")
  case "photo"
      session("type")="photo"
	  session("seq")=Request("seq")
      session("cust_area")=left(Request.QueryString("cust_area"),1)& gdept
      session("draw_file")=Request("draw_file")
      '93/03/08_jessica修改
      session("form_name")=Request("form_name")
      '93/03/08_end
      cont="圖檔上傳"
  case "doc"
      session("type")="doc"
	  session("folder_name")=Request("folder_name")	'檔案目錄所在,會在cust_area之下建立
	  session("prefix_name")=Request("prefix_name")	'檔案的前置名稱:例如: filename=abc.jpg, if prefix="123" , then filename=123_abc.jpg,用於區隔同目錄底下,不同的檔案名稱
      session("cust_area")=left(Request.QueryString("cust_area"),1) & gdept
      session("draw_file")=Request("draw_file")	'原檔案路徑(存在於server上的D:\data\document)
      session("form_name")=Request("form_name")
      session("size_name")=Request("size_name")
      session("file_name")=Request("file_name")
      session("source_name")=Request("source_name")
      session("btnname")=Request("btnname")
      session("nfilename")=Request("nfilename")
      session("doc_in_date")=Request("in_date")
      session("doc_in_scode")=Request("in_scode")
      cont="檔案上傳"
  case "Ext_photo"
      session("type")="photo"
      session("seq")=Request("seq")
      session("cust_area")=left(Request.QueryString("cust_area"),1)&"PE"
      session("draw_file")=Request("draw_file")
      '93/03/08_jessica修改
      session("form_name")=Request("form_name")
      '93/03/08_end
      cont="圖檔上傳"
  case else
      session("type")=""
      response.write "<html><head><title>RE?!1ORE!?DAo3?μoμ!!C</title></head><body bgcolor=#ffffff><br><br><p><center>RE?!1ORE!?DAo3?μoμ!!C"
      response.write "<form><input type=button value=Ao3?μoμ! onclick=""javascript:parent.close()""></form></center></body></html>"
      Response.End
end select

%>
<body bgcolor="#FFFFFF">
<p align="center"><big><font face="標楷體" color="#004000"><strong><big><big><%=cont%></big></big></strong></font></big></p>

<center>
  <form name="AttachForm" action="upload_winact_file.asp" method="Post" enctype="multipart/form-data">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr><td>
        <div align="left">
          　上傳檔案到本資料欄位:
          <br>
          　<input type="file" name="theFile" size="25">
          　<input type="hidden" name="hidFile" size="25">
          　<input type="hidden" name="hidoverwrite" size="25">
          　<input type="hidden" name="nfilename" size="<%=Request("nfilename")%>">
          　<input type="hidden" name="tablename" size="<%=Request("tablename")%>">

          <br>&nbsp;
<span style="display:none">          
<font size="2" color="red">
<input type="checkbox" id="chkoverwrite" name="chkoverwrite">覆蓋已存在的檔案<br></font></span>
          <br>
          <table width="95%" border="0">
            <tr> 
              <td>
<!--<font color="red"><strong>[請注意]</strong>上傳檔案應先取得著作權人的書面同意，不得有侵犯他人著作權的行為!</font><br><br>-->
<font size="2" color="#009900">使用方式：</font><br>
<div align="center"><center>

<table border="0" width="100%">
  <tr>
    <td width="9%" align="right" valign="top"><font size="2" color="black">◎</td>
    <td width="91%"><font size="2" color="black">
    欲上傳檔案至本欄位，請點選上方之『瀏覽』按鈕後會出現一個『選擇檔案』小視窗，然後請選擇您電腦中欲上傳之檔案。</font>
    </td>
  </tr>
<!--  <tr>
    <td width="9%" align="right" valign="top"><font size="2" color="black">◎</td>
    <td width="91%"><font size="2" color="red">檔案名不可以有中文字。</td>
  </tr>-->
</table>
  </center></div>
</td></tr>
</table>
        </div>
</td></tr>
<tr><td align="center">
<input type="button" value="上傳" onclick="AttachFile()" id="button1" name="button1" class="cbutton">
<input type="button" value="關閉視窗" onclick="javascript:parent.close()" id="button2" name="button2" class="cbutton">
</td></tr>
  </table>
</form>
</center></body>
</html>
