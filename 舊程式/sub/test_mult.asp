<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<title>test</title>
</head>    
<body>
<form method="POST" name="reg" action="">
<input type=text id="EPctcountry" name="EPctcountry" size=50>
<input type=button id=button1 name=button1 value="..." onclick="chomul('EPctcountry')">
</form>
</body>
</html>
<script language=vbscript>
function chomul(p1)
'http://web01/brp/sub/test_mult.asp
	<%session("mult") = session("ODBCDSN")%>
	isql = "select coun_code,coun_c from country order by coun_code"
	window.open "Client_mult.asp?isql=" & isql & "&field1=" & p1, "myWindowOne", "width=550 height=400 top=100 left=100 toolbar=no, menubar=no, location=no, directories=no resizeable=no status=no scrollbar=no"
end function
</script>