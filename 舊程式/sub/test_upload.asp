<PRE class=clsCode style="DISPLAY: block" AUTOHILITE="1"><HTML>
<HEAD>
<TITLE>File Upload Control</TITLE>
</HEAD>

<BODY LEFTMARGIN=20 TOPMARGIN=20 BGCOLOR=#FFFFFF TEXT=#000000 LINK=#FF0066 VLINK=#330099 ALINK=#000000 language="VBS" onload="InitializeControl">

<FONT FACE="ARIAL" SIZE=2>

<CENTER>
<P><FONT SIZE=5><B>Welcome</B></FONT>
</CENTER>

<H3>File Upload</H3>
<P>
You can upload your HTML content using this control.
<OBJECT classid="clsid:886E7BF0-C867-11CF-B1AE-00AA00A3F2C3" id=FlUpl1 width=100 height=100
    align=textmiddle color=blue codebase=http://my_server/flupl/FlUpl.cab#Version=6,0,86,0>
</OBJECT> 
You can drag and drop files onto it. Or double-click it to get 
  a FileOpen dialog box.

<SCRIPT LANGUAGE="VBS">
Sub InitializeControl
        FlUpl1.CreateTempBinding "http://my_server", 
    "{8B14B770-748C-11D0-A309-00C04FD7CFC5}"
End Sub
</SCRIPT>

<P ALIGN=CENTER><FONT SIZE=2><B>For more 
  information on VBScript, visit the <A TARGET="_top" 
  HREF="/vbscript/default.htm">VBScript site.</A></B>
  </FONT>

</BODY>
</HTML>
</PRE>