<%'http://web01/brp/sub/test_layer.asp%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<TITLE>Layer</TITLE>
<BODY bgcolor="#FFFFFF" topMargin=2>
<form name=reg>
<table>
<tr><td><input type=button name=button3 value="Enter" onclick="onlayer('1')"></td></tr>
<tr><td></td></tr>
<tr><td><input type=button name=button4 value="Enter" onclick="onlayer2('2')"></td></tr>
</table>
<br><br><br><br><br><br>
<input type=button name=button1 value="Enter" onclick="onlayer('1')" scrolltop=0 scrollleft=0><br><br>
<input type=button name=button2 value="Enter" onclick="onlayer('2')">
<input type=text value="Aaa">
<div id=layer1 name=layer1 align="center" style="display:none;position:absolute">
test test
</div>
</form>
</BODY>
</HTML>
<script language="vbscript">
function onlayer(p1)
	document.all.layer1.style.display = ""
'	execute "document.all.layer1.style.top = document.all.button" & p1 & ".style.top"
'	execute "document.all.layer1.style.left = document.all.button" & p1 & ".style.left+document.all.button" & p1 & ".style.width"
	execute "document.all.layer1.style.top = document.all.button1.style.top-10"
	execute "document.all.layer1.style.left = document.all.button1.style.left+document.all.button1.style.width"
end function
function onlayer2(p1)
	document.all.layer1.style.display = ""
	execute "document.all.layer1.style.top = document.all.button2.style.top"
	execute "document.all.layer1.style.left = document.all.button2.style.left+document.all.button2.style.width"
end function
</script>
