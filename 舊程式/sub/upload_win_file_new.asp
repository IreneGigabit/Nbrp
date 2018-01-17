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
		alert "請輸入要上傳的檔案名稱，或使用瀏覽來選擇檔案。"
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
   case "apcust_file","custdb_file","db_file","custresp_file","brdb_file"
      '2012/5/2增加，custdb_file=對催帳客函 db_file=請款單 custresp_file=客戶對催回應文件
      '2015/11/16 apcust_file 契約書、委任書
	  '2016/11/22增加brdb_file=英文invoice
      session("type")=request.querystring("type")
	  session("folder_name")=Request("folder_name")	'檔案目錄所在,會在cust_area之下建立
	  session("prefix_name")=Request("prefix_name")	'檔案的前置名稱:例如: filename=abc.jpg, if prefix="123" , then filename=123_abc.jpg,用於區隔同目錄底下,不同的檔案名稱
      session("cust_area")=left(Request.QueryString("cust_area"),1) & gdept
      session("draw_file")=Request("draw_file")	'原檔案路徑(存在於server上的D:\data\document)
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
      session("ar_no")=request("ar_no")	'for 英文invoice
      session("qs_dept")=request("qs_dept")	'for 英文invoice
      session("draw_name")=request("draw_name")	'for 英文invoice
      cont="檔案上傳"    
      if request.querystring("type")="custdb_file" then
         cont="對催帳客函檔案上傳"    
      elseif request.querystring("type")="db_file" then
         cont="請款單檔案上傳"  
      elseif request.querystring("type")="custresp_file" then
         cont="對催帳客戶回應檔案上傳"
      elseif request.querystring("type")="apcust_file" then
        session("db_file_flag") = ""
      elseif request.querystring("type")="brdb_file" then
        cont="英文Invoice檔案上傳"
        'session("db_file_flag") = ""
      end if
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
  <form id="AttachForm" name="AttachForm" action="upload_winact_file_new.asp" method="Post" enctype="multipart/form-data">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr><td>
        <div align="left">
          　上傳檔案到本資料欄位:
          <br>
          　<input type="file" id="theFile" name="theFile" size="25">
          　<input type="hidden" id="hidFile" name="hidFile" size="25">
          　<input type="hidden" id="hidoverwrite" name="hidoverwrite" size="25">
          　<input type="hidden" id="nfilename" name="nfilename" size="<%=Request("nfilename")%>">
			<input type="hidden" id="tablename" name="tablename" size="<%=Request("tablename")%>">
			<input type="hidden" id="testshow" name="testshow" size="<%=Request("folder_name")%>">
          <br>&nbsp;
<span style="display:none">          
<font size="2" color="red">
<input type="checkbox" id="chkoverwrite" name="chkoverwrite">覆蓋已存在的檔案<br></font></span>
          <br>
          <table width="95%" border="0">
            <tr> 
              <td>
<!--<font color="red"><strong>[請注意]</strong>上傳檔案應先取得著作權人的書面同意，不得有侵犯他人著作權的行為!</font><br><br>-->
<div align="center"><center>

<table border="0" width="100%">
    <tr>
        <td align="left" colspan=2  style="font:14px;color:#009900">使用方式：</td>
    </tr>
    <tr><td width="4%" align="right" valign="top">◎</td>
        <td width="96%" style="font:14px;color:black">
        欲上傳檔案至本欄位，請點選上方之『瀏覽』按鈕後會出現一個『選擇檔案』小視窗，然後請選擇您電腦中欲上傳之檔案。
        </td>
    </tr>
    <tr><td width="4%" align="right" valign="top">◎</td>
        <td width="96%" style="font:14px;color:red">
        若需上傳Word檔，請於儲存檔案時，檔案類型選擇「Word 97-2003文件(*.doc)」，以免日後檢視無法開啟。
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
