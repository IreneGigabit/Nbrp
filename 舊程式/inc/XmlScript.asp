<script language=vbs>
sub xml_select(url, sel, key)
	dim xmldoc, root, x
	dim html, head, body
	
	html = sel.outerHTML
	head = mid(html, 1, instr(1, html, "<OPTION") - 1)
	body = "<OPTION VALUE="""" STYLE=""COLOR:BLUE"">請選擇</OPTION>"
	
	set xmldoc = CreateObject("Microsoft.XMLDOM")
	xmldoc.async = false
	xmldoc.validateOnParse = true
'msgbox url
	if xmldoc.load (url) then
		set root = xmldoc.documentElement		
		for each x in root.childNodes
			if x.childNodes.item(0).text = key then
				body = body & "<OPTION VALUE=""" & x.childNodes.item(0).text & """ selected>" & x.childNodes.item(1).text & "</OPTION>"
			else
				body = body & "<OPTION VALUE=""" & x.childNodes.item(0).text & """>" & x.childNodes.item(1).text & "</OPTION>"
			end if
		next
	end if
	
	set xmldoc = nothing
	set root = nothing
	sel.outerHTML = head & body & "</SELECT>"
'msgbox sel.outerHTML
end sub

sub xml_select2(url, sel, key)
	
	dim xmldoc, root, x
	dim html, head, body
	html = sel.outerHTML
	sp = instr(1, html, "name=") + 5 	
	head1 = mid(html, sp )	
	sp2 = instr(1, head1, ">")
	sp3 = instr(1, head1, " ")
		
	if sp2 = 0 then
		szname = mid(head1,1,sp3 - 1)
	end if
	if sp3 = 0 then
		szname = mid(head1,1,sp2 - 1)
	end if
	if sp2 > 0 and sp3 > 0 then
		if sp2 < sp3 then
			szname = mid(head1,1,sp2 - 1)
		else
			szname = mid(head1,1,sp3 - 1)
		end if
	end if
	head = "<select Name=" & szname & " size=1>"
	body = "<OPTION VALUE="""" STYLE=""COLOR:BLUE"">請選擇</OPTION>"
	
	set xmldoc = CreateObject("Microsoft.XMLDOM")
	xmldoc.async = false
	xmldoc.validateOnParse = true
	if xmldoc.load (url) then
		set root = xmldoc.documentElement		
		for each x in root.childNodes
			if x.childNodes.item(0).text = key then
				body = body & "<OPTION VALUE=""" & x.childNodes.item(0).text & """ selected>" & x.childNodes.item(1).text & "</OPTION>"
			else
				body = body & "<OPTION VALUE=""" & x.childNodes.item(0).text & """>" & x.childNodes.item(1).text & "</OPTION>"
			end if
		next
	end if
	
	set xmldoc = nothing
	set root = nothing
	
	sel.outerHTML = head & body & "</SELECT>"
end sub
sub xml_select3(url, sel,szname, key)
	dim xmldoc, root, x
	dim html, head, body
	head = "<select Name=" & szname & " size=1>"
	body = "<OPTION VALUE="""" STYLE=""COLOR:BLUE"">請選擇</OPTION>"
	set xmldoc = CreateObject("Microsoft.XMLDOM")
	xmldoc.async = false
	xmldoc.validateOnParse = true
	if xmldoc.load (url) then
		set root = xmldoc.documentElement		
		for each x in root.childNodes
			if x.childNodes.item(0).text = key then
				body = body & "<OPTION VALUE=""" & x.childNodes.item(0).text & """ selected>" & x.childNodes.item(1).text & "</OPTION>"
			else
				body = body & "<OPTION VALUE=""" & x.childNodes.item(0).text & """>" & x.childNodes.item(1).text & "</OPTION>"
			end if
		next
	end if
	
	set xmldoc = nothing
	set root = nothing
	kname=head & body & "</SELECT>"
	execute "reg."&szname&".outerHTML = kname"
end sub
</script>