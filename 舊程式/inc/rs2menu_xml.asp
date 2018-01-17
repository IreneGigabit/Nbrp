<%
dim img(13)
dim ref

img(0) = "<img border=0 align=absmiddle src=images/tv_dots.gif width=16 height=20>"
img(1) = "<img border=0 align=absmiddle src=images/tv_dotsb.gif width=16 height=20>"
img(2) = "<img border=0 align=absmiddle src=images/tv_plusdots.gif width=16 height=20>"
img(3) = "<img border=0 align=absmiddle src=images/tv_plusdotsb.gif width=16 height=20>"
img(4) = "<img border=0 align=absmiddle src=images/tv_minusdots.gif width=16 height=20>"
img(5) = "<img border=0 align=absmiddle src=images/tv_minusdotsb.gif width=16 height=20>"
img(6) = "<img border=0 align=absmiddle src=images/tv_dotsl.gif width=16 height=20>"
img(7) = "<img border=0 align=absmiddle src=images/tv_space.gif width=16 height=20>"
img(8) = "<img border=0 align=absmiddle src=images/Clsdfold.gif width=16 height=16>"
img(9) = "<img border=0 align=absmiddle src=images/Openfold.gif width=16 height=16>"
img(10) = "<img border=0 align=absmiddle src=images/tv_plusdotsbt.gif width=16 height=20>"
img(11) = "<img border=0 align=absmiddle src=images/tv_minusdotsbt.gif width=16 height=20>"
img(12) = "<img border=0 align=absmiddle src=images/user.gif width=16 height=16>"

sub rs2menu_xml(sql, xmldoc)
	' 欄位順序:
	'  rs(0): nodeID
	'  rs(1): parentID, xzzz 表示 root
	'  rs(2): name
	'  rs(3): href
	'  rs(4): kind
	' 注意事項:
	'  1. 須依階層排列, 一階、二階、三階 ...
	'  2. rs(0)、rs(1) 自動加 'x' 以防止數字時錯誤
	
	dim cn, rs, root, node, parent, name, kind, href

	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	cn.Open Session("ODBCDSN")
	rs.Open sql, cn, 0, 1

	if rs.EOF then
		set root = xmldoc.createElement("menu")
		xmldoc.appendChild root
		set node = xmldoc.createElement("item")
		root.appendChild(node)
		node.setAttribute "name", "目前還沒有任何資料!!"
		node.setAttribute "href", ""
		node.setAttribute "kind", "d"
	else
		set root = xmldoc.createElement("menu")
		xmldoc.appendChild root
		do while not rs.EOF		
			node = "x" + trim(rs(0))
			parent = "x" + trim(rs(1))
			name = trim(rs(2))
			href = trim(rs(3))
			kind = trim(rs(4))
			if parent = "xzzz" then
				parent = "root"
			end if
			execute("set " & node & " = xmldoc.createElement(""item"")")
			execute(parent & ".appendChild(" & node & ")")
			execute(node & ".setAttribute ""name"",""" & name & """")
			execute(node & ".setAttribute ""href"",""" & href & """")
			execute(node & ".setAttribute ""kind"",""" & kind & """")
			rs.MoveNext
		loop
	end if

	rs.Close
	cn.Close

	set rs = nothing
	set cn = nothing	
end sub

sub rs2menu_xml_1(sql, xmldoc)
	' 欄位順序:
	'  rs(0): nodeID
	'  rs(1): parentID, xzzz 表示 root
	'  rs(2): name
	'  rs(3): href
	' 注意事項:
	'  1. 須依階層排列, 一階、二階、三階 ...
	'  2. rs(0)、rs(1) 自動加 'x' 以防止數字時錯誤
	
	dim cn, rs, root, node, parent, name, href

	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	cn.Open Session("ODBCDSN")
'Response.Write sql & "<BR>"
	rs.Open sql, cn, 0, 1
'Response.Write rs.EOF & "<BR>"
	if rs.EOF then
		set root = xmldoc.createElement("menu")
		xmldoc.appendChild root
		set node = xmldoc.createElement("item")
		root.appendChild(node)
		node.setAttribute "name", "目前還沒有任何資料!!"
		node.setAttribute "href", ""
	else
		set root = xmldoc.createElement("menu")
		xmldoc.appendChild root
'i=1
		do while not rs.EOF
'Response.Write i & "<BR>"
			node = "x" + trim(rs(0))
			parent = "x" + trim(rs(1))
			name = trim(rs(2))
			href = trim(rs(3))
			if parent = "xzzz" then
				parent = "root"
			end if
			a = "set " & node & " = xmldoc.createElement(""item"")"
'Response.Write a & "<BR>"
			execute a
'			execute("set " & node & " = xmldoc.createElement(""item"")")
'Response.Write "ERR.number: " & ERR.number & ERR.description & "<BR>"
			a = parent & ".appendChild(" & node & ")"
'Response.Write a & "<BR>"
			execute a
'			execute(parent & ".appendChild(" & node & ")")
'Response.Write "ERR.number: " & ERR.number & ERR.description & "<BR>"
			a = node & ".setAttribute ""name"",""" & name & """"
'Response.Write a & "<BR>"
			execute a
'			execute(node & ".setAttribute ""name"",""" & name & """")
'Response.Write "ERR.number: " & ERR.number & ERR.description & "<BR>"
			a = node & ".setAttribute ""href"",""" & href & """"
'Response.Write a & "<BR>"
			execute a
'			execute(node & ".setAttribute ""href"",""" & href & """")
'Response.Write "ERR.number: " & ERR.number & ERR.description & "<BR>"
'Response.End
'i=i+1
			rs.MoveNext
		loop
	end if

	rs.Close
	cn.Close

	set rs = nothing
	set cn = nothing	
end sub

sub displayMenuTree(nodes, depth, last)
	dim i, j
	dim name, href, kind, childs
	
	for i = 0 to nodes.length - 1
		if i = nodes.length - 1 then
			call StrReplace(last, depth+1, depth+2, "1")
		else
			call StrReplace(last, depth+1, depth+2, "0")
		end if
		name = nodes.item(i).getAttribute("name")
		href = nodes.item(i).getAttribute("href")
		kind = nodes.item(i).getAttribute("kind")
		if nodes.item(i).getElementsByTagName("item").length = 0 then
			' 第一層選項
			Response.Write "<div>"
			for j = 1 to depth
				if mid(last, j, 1) = "0" then
					Response.Write img(6)
				else
					Response.Write img(7)
				end if
			next
			if mid(last, j, 1) = "0" then
				Response.Write img(0)
			else
				Response.Write img(1)
			end if
			if kind="d" then
				Response.Write img(8)
			else
				Response.Write img(12)
			end if
			if href <> "" then
				Response.Write "&nbsp;<A target='content' href='" & href & "')>" & name & "</A>"
			else
				Response.Write "&nbsp;<SPAN class=error>" & name & "</SPAN>"
			end if
			Response.Write "</div>" & vbcrlf
		else
			' 擁有子選項的選項
			ref = ref + 1
			Response.Write "<div>"
			
			' 加
			Response.Write "<span id='p" & ref & "' style='cursor:hand' onClick='show(""" & ref & """)'>"
			for j = 1 to depth							
				if mid(last, j, 1) = "0" then
					Response.Write img(6)
				else
					Response.Write img(7)
				end if
			next
			if depth = 0 and j = 1 then
				Response.Write img(10)
			else
				if mid(last, j, 1) = "0" then
					Response.Write img(2)
				else
					Response.Write img(3)
				end if
			end if
			if kind="d" then
				Response.Write img(8)
			else
				Response.Write img(12)
			end if
			Response.Write "</span>"
			
			' 減
			Response.Write "<span id='m" & ref & "' style='cursor:hand;display:none' onClick='hide(""" & ref & """)'>"
			for j = 1 to depth
				if mid(last, j, 1) = "0" then
					Response.Write img(6)
				else
					Response.Write img(7)
				end if
			next
			if depth = 0 and j = 1 then
				Response.Write img(11)
			else
				if mid(last, j, 1) = "0" then
					Response.Write img(4)
				else
					Response.Write img(5)
				end if
			end if
			if kind="d" then
				Response.Write img(9)
			else
				Response.Write img(12)
			end if
			Response.Write "</span>"
			
			if href <> "" then
				Response.Write "&nbsp;<A target='content' href='" & href & "')>" & name & "</A>" & vbcrlf
			else
				Response.Write "&nbsp;<SPAN class=error>" & name & "</SPAN>"			
			end if
			
			' 子選項
			Response.Write "<div id='s" & ref & "' style='display:none'>" & vbcrlf
			
			set childs = nodes.item(i).selectNodes("item")
            call displayMenuTree(childs, depth + 1, last) 
            
            Response.Write "</div></div>" & vbcrlf
		end if
	next
end sub

sub displayMenuTree_1(nodes, depth, last)
	dim i, j
	dim name, href, childs
	
	for i = 0 to nodes.length - 1
		if i = nodes.length - 1 then
			call StrReplace(last, depth+1, depth+2, "1")
		else
			call StrReplace(last, depth+1, depth+2, "0")
		end if
		name = nodes.item(i).getAttribute("name")
		href = nodes.item(i).getAttribute("href")
		if nodes.item(i).getElementsByTagName("item").length = 0 then
			' 第一層選項
			Response.Write "<div>"
			for j = 1 to depth
				if mid(last, j, 1) = "0" then
					Response.Write img(6)
				else
					Response.Write img(7)
				end if
			next
			if mid(last, j, 1) = "0" then
				Response.Write img(0)
			else
				Response.Write img(1)
			end if
			Response.Write img(8)
			if href <> "" then
				Response.Write "&nbsp;<A target='content' href='" & href & "')>" & name & "</A>"
			else
				Response.Write "&nbsp;<SPAN class=error>" & name & "</SPAN>"
			end if
			Response.Write "</div>" & vbcrlf
		else
			' 擁有子選項的選項
			ref = ref + 1
			Response.Write "<div>"
			
			' 加
			Response.Write "<span id='p" & ref & "' style='cursor:hand' onClick='show(""" & ref & """)'>"
			for j = 1 to depth							
				if mid(last, j, 1) = "0" then
					Response.Write img(6)
				else
					Response.Write img(7)
				end if
			next
			if depth = 0 and j = 1 then
				Response.Write img(10)
			else
				if mid(last, j, 1) = "0" then
					Response.Write img(2)
				else
					Response.Write img(3)
				end if
			end if
			Response.Write img(8)
			Response.Write "</span>"
			
			' 減
			Response.Write "<span id='m" & ref & "' style='cursor:hand;display:none' onClick='hide(""" & ref & """)'>"
			for j = 1 to depth
				if mid(last, j, 1) = "0" then
					Response.Write img(6)
				else
					Response.Write img(7)
				end if
			next
			if depth = 0 and j = 1 then
				Response.Write img(11)
			else
				if mid(last, j, 1) = "0" then
					Response.Write img(4)
				else
					Response.Write img(5)
				end if
			end if
			Response.Write img(9)
			Response.Write "</span>"
			
			if href <> "" then
				Response.Write "&nbsp;<A target='content' href='" & href & "')>" & name & "</A>" & vbcrlf
			else
				Response.Write "&nbsp;<SPAN class=error>" & name & "</SPAN>"			
			end if
			
			' 子選項
			Response.Write "<div id='s" & ref & "' style='display:none'>" & vbcrlf
			
			set childs = nodes.item(i).selectNodes("item")
            call displayMenuTree_1(childs, depth + 1, last) 
            
            Response.Write "</div></div>" & vbcrlf
		end if
	next
end sub

sub StrReplace(s1, pos1, pos2, s2)
	s1 = mid(s1, 1, pos1 - 1) & _
		 s2 & _
		 mid(s1, pos2 + 1)	
end sub
%>
<STYLE>
A:hover
{
    FONT-SIZE: 9pt;
    COLOR: red;
    FONT-FAMILY: Verdana, Arial;
}
A
{
    FONT-SIZE: 9pt;
    COLOR: #000066;
    FONT-FAMILY: Verdana, Arial;
}
.error
{
    FONT-SIZE: 9pt;
    COLOR: red;
    FONT-FAMILY: Verdana, Arial;
}
</STYLE>
<SCRIPT language="VBScript">
' 顯示子選項
function show(ele)
    document.all("s" + ele).style.display = ""
    document.all("p" + ele).style.display = "none"
    document.all("m" + ele).style.display = ""
end function
' 隱藏子選項
function hide(ele)
    document.all("s" + ele).style.display = "none"
    document.all("p" + ele).style.display = ""
    document.all("m" + ele).style.display = "none"
end function
</SCRIPT>
