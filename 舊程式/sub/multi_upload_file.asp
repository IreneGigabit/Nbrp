<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>多檔上傳</title>
<link rel="stylesheet" type="text/css" href="../js/swfupload/uuuu.css" />
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css" />
<script type="text/javascript" src="../js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="../js/swfupload/swfupload.js"></script>
<script type="text/javascript" src="../js/swfupload/swfupload.queue.js"></script>
<script type="text/javascript" src="../js/swfupload/fileprogress.js"></script>
<script type="text/javascript" src="../js/swfupload/handlers.js"></script>
<%
uploadfield = "attach"
prgid = request("prgid")
seqdept = request("seqdept") 'P內專、PE出專
seq = request("seq")
seq1 = request("seq1")
step_grade = request("step_grade")
job_sqlno = request("job_sqlno")
upfolder = request("upfolder")
'attach_no = request("attach_no")
screen_no = request("screen_no")
screen_no = screen_no + 1
session("screen_no") = screen_no
attach_tablename = request("attach_tablename")
temptable = request("temptable")

fseq = session("se_branch") & session("dept") & seq
if seq1<>"_" then fseq = fseq &"-"& seq1

'response.Write seq & "<BR>"
'response.Write seq1 & "<BR>"
'response.Write step_grade & "<BR>"
'response.Write job_sqlno & "<BR>"
'response.Write "attach_no="& attach_no & "<BR>"
%>
<script type="text/javascript" language="javascript">
    //必須用javascript寫
    var upload2;
    var screen_no;

    $(function() {
		upload2 = new SWFUpload({
			// Backend Settings
		    //處理實際檔案上傳
		    upload_url: "../sub/UpLoadFile.asp",  
		    //欲傳入的參數
		    post_params: { "prgid": "<%=prgid%>", "uploadfield": "<%=uploadfield%>", "seqdept": "<%=seqdept%>", "seq": "<%=seq%>", "seq1": "<%=seq1%>", "step_grade": "<%=step_grade%>", "job_sqlno": "<%=job_sqlno%>", "upfolder": "<%=upfolder%>", "attach_tablename": "<%=attach_tablename%>", "attach_no": "<%=session("attach_no")%>", "temptable": "<%=temptable%>" },

			// File Upload Settings
		    //file_size_limit: "204800", // 100MB
		    file_size_limit: "81920", // 40MB
			file_types : "*.*",                    //顯示什麼副檔名的檔案
			file_types_description : "All Files",  //可選擇的檔案類型
			file_upload_limit: "20",  //最大文件上傳大小
			file_queue_limit: "0",  // Zero means unlimited

			// 檔案處理函式 Event Handler Settings (all my handlers are in the Handler.js file)
			file_dialog_start_handler : fileDialogStart,
			file_queued_handler : fileQueued,
			file_queue_error_handler : fileQueueError,
			file_dialog_complete_handler : fileDialogComplete,
//			upload_start_handler : uploadStart,
			upload_start_handler : uploadStart1,
			upload_progress_handler : uploadProgress,
			upload_error_handler : uploadError,
//			upload_success_handler : uploadSuccess,
			upload_success_handler : uploadSuccess1,
			upload_complete_handler : uploadComplete,

			// 按鈕設定 Button Settings
			button_placeholder_id : "spanButtonPlaceholder2",
			//button_image_url : "../js/swfupload/XPButtonUploadText_61x22.png", '原圖
			//button_image_url: "../js/swfupload/XPButtonUploadText.png",
			button_image_url: "../js/swfupload/XPButtonUploadText1.jpg", //---無字圖
			button_text : "<span class='thebutton'>&nbsp;瀏　覽&nbsp;</span>",
			button_text_style: ".thebutton { font-size: 16; BORDER-TOP: medium none; }",			
			button_width: 61,
			button_height: 22,
			button_cursor: SWFUpload.CURSOR.HAND,
			
			// 要加入的元件 Flash Settings
			flash_url : "../js/swfupload/swfupload.swf",

			custom_settings : {
				progressTarget : "fsUploadProgress2",
				cancelButtonId : "btnCancel2"
			}//,
			// Debug Settings
			//debug: true
		});
    });
	
	var fileAllSize = 0;
	function uploadStart1(f) {
	    fileAllSize += f.size;
	    //alert("Startimg...... " + fileAllSize +  " ...!!!");
	    return true;
	}

	function uploadSuccess1(f, m, rep) {
	    $("#repmsg").html("Good: " + $("#repmsg").html() + " " + m);
	    var args = m.split("#@#");
	    //alert(args[7]);
	    //alert(args[2]);
	    //alert(args[3]);
	    if (args[0] == "1") {
	        //將轉回的資料丟回畫面
	        //6:attach_no，7:檔案名稱，8:share folder完整路徑，9:虛擬完整路徑，10:原始檔名，11:檔案大小
	        SetpValue(args[6], args[7], args[8], args[9], args[10], args[11]); 
	        
	        //screen_no = parseInt(screen_no) + 1;
	        <%session("screen_no") = session("screen_no") + 1%>
	        //alert("<%=session("attach_no")%>");
	    } else {
			this.uploadError(f,args[0],args[1]);
	    }
	    try {	        
	        var progress = new FileProgress(f, this.customSettings.progressTarget);
	        progress.setComplete();
	        progress.setStatus("上傳成功");
	        progress.toggleCancel(false);
	    } catch (ex) {
	        this.debug(ex);
	    }
	}
	
	function SetpValue(pattach_no,pattach_name, psPath, pattach_path, psource_name, pattach_size) {
	    //檔案名稱，虛擬完整路徑，原始檔名，檔案大小，attach_no
	    var pvalue = pattach_name+"#@#"+pattach_path+"#@#"+psource_name+"#@#"+pattach_size+"#@#"+pattach_no
	    //alert(pvalue);
	    window.opener.AddFileattach(pvalue);
	}
</script>
</head>
<body>
<form id="reg" name="reg" method="post" action="">
<table cellspacing="1" cellpadding="0" width="100%" border="0">
<tr>
    <td width="50%" nowrap="nowrap" class="FormName">【<%=prgid%>&nbsp;多檔上傳】<%=fseq%></td>
    <td width="50%" nowrap="nowrap" align="right" class="FormName">
        <font style="cursor: hand;color:darkblue" onmouseover="vbs:me.style.color='red'" onmouseout="vbs:me.style.color='darkblue'"  onclick="vbscript:window.close">[關閉視窗]</font>
    </td>
</tr>
<tr>
    <td colspan=2>
    <!--多檔上傳 Begin-->
    <input type="hidden" name="seq" value="<%=seq%>" />
    <input type="hidden" name="seq1" value="<%=seq1%>" />
    <input type="hidden" name="step_grade" value="<%=step_grade%>" />
    <input type="hidden" name="job_sqlno" value="<%=job_sqlno%>" />
    <input type="hidden" name="TempUpFile1" value="" />
    <br />
    <div class="fieldset flash" id="fsUploadProgress2"><span class="legend">多檔上傳</span></div>
    <div style="padding-left: 5px;">
    <span id="spanButtonPlaceholder2"></span>
    <!--<input id="spanButtonPlaceholder2" type="button" value="瀏　覽"/>-->
    &nbsp;&nbsp;<input id="btnCancel2" type="button" value="取消上傳" onclick="cancelQueue(upload2);" disabled="disabled" style="margin-left: 2px; height: 28px; font-size: 8pt;" />
    </div>
    <!--多檔上傳 End-->
    </td>
</tr>
</table>
</form>
</body>
</html>
