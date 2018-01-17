<%
'http://sin09/brp/sub/data/ap_exp_expcount.asp
'抓申請人案件統計
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open session("btbrtdb")
'Set cnn = Server.CreateObject("ADODB.Connection")
'cnn.Open session("ODBCDSN")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../inc/setstyle.css">
<title></title>
</head>
<body>
申請人案件統計<br>
<%
Set rsi = Server.CreateObject("ADODB.RecordSet")
isql = "select year(in_date),count(*) from exp"
isql = isql & " where in_date>='2006/1/1' and in_date<='2007/12/31'"
isql = isql & " and seq is not null and seq<>'0'"
isql = isql & " group by year(in_date)"
rsi.Open isql,conn,1,1
apsum = 0
while not rsi.EOF 
'	Response.Write rsi(0) &":"& rsi(1) &"筆<br>"
	apsum = apsum + cdbl(rsi(1))
	rsi.MoveNext 
wend
rsi.Close 
%>
<form name="reg" method="POST" action="">
<table border="1" class="greentable" cellspacing="1" cellpadding="1" width="80%">
	<TR class=greenths>
		<TD align=left width=20%>出專&nbsp;<%=session("se_branch")%>&nbsp;(<%=apsum%>)</TD>
		<TD align=center width=12%>公司</TD>
		<TD align=center width=12%>個人</TD>
		<TD align=center width=12%>其他</TD>
		<TD align=center width=15%>合計</TD>
		<TD align=center width=14%>有問題資料</TD>
		<TD align=center width=15%>總計</TD>
	</TR>
	<TR>
		<TD class=greentext align=right>2006：</TD>
		<TD class=greendata align=left><INPUT type=text name="Aap2006" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="Bap2006" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="cap2006" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="ap2006sum" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="ap2006err" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="sum2006" size="6" value=0 readonly class=gsedit></td>
	</TR>
	<TR>
		<TD class=greentext align=right>2007：</TD>
		<TD class=greendata align=left><INPUT type=text name="Aap2007" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="Bap2007" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="Cap2007" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="ap2007sum" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="ap2007err" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="sum2007" size="6" value=0 readonly class=gsedit></td>
	</TR>
	<TR>
		<TD class=greentext align=right>合計：</TD>
		<TD class=greendata align=left><INPUT type=text name="Aapsum" size="6" value=0 readonly class=gsedit></td>
		<TD class=greendata align=left><INPUT type=text name="Bapsum" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="Capsum" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="apsum" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="aperr" size="6" value=0 readonly class=gsedit></TD>
		<TD class=greendata align=left><INPUT type=text name="totalsum" size="6" value=0 readonly class=gsedit></td>
	</TR>
	<TR>
		<TD class=greendata align=left colspan=7>
			<textarea name="remark" cols=80 rows=20>
			</textarea>
		</TD>
	</TR>
	
</table>
</td></tr>  
</table> 
</form>
<%
'select * From exp where convert(varchar,seq)+seq1 not in (select convert(varchar,seq)+seq1 From ap_exp )
'and in_date>='2006/1/1' and in_date<='2007/12/31' and seq is not null and seq<>'0'
%>
<script Language="VBScript">
function window_onload()
	msgremark = ""
	err = 0
	fsql = "select seq,seq1,in_date from exp"
	fsql = fsql & " where in_date>='2006/1/1' and in_date<'2008/1/1'"
	fsql = fsql & " and seq is not null and seq<>'0'"
	url = "../../xml/XmlGetSqlData.asp?searchsql=" & fsql
	'window.open url
	set xmldoc = CreateObject("Microsoft.XMLDOM")
	xmldoc.async = false
	xmldoc.validateOnParse = true
	if xmldoc.load (url) then
		Set root = xmldoc.documentElement
		For Each xi In root.childNodes
			sql = "select b.apclass,b.apcust_no,b.apsqlno,b.ap_cname1,b.ap_cname2 from ap_exp a,apcust b"
			sql = sql & " where a.apsqlno=b.apsqlno and a.seq="& xi.childNodes.item(1).text
			sql = sql & " and seq1='"& xi.childNodes.item(2).text &"'"
			sql = sql & " order by a.sqlno"
			url = "../../xml/xmlgetsqldata.asp?searchsql=" & sql
			'window.open url
			set xmldocs = CreateObject("Microsoft.XMLDOM")
			xmldocs.async = false
			xmldocs.validateOnParse = true
			if xmldocs.load (url) then
			'	Set roots = xmldocs.documentElement
			'	For Each xi In root.childNodes
			'		
			'	next
			'	set roots = nothing
				if xmldocs.selectSingleNode("//xhead/Found").text = "Y" then
					if trim(xmldocs.selectSingleNode("//xhead/apclass").text)<>empty then
						if left(xmldocs.selectSingleNode("//xhead/apclass").text,1)="A" then
							execute "reg.Aap"& year(xi.childNodes.item(3).text) &".value = reg.Aap"& year(xi.childNodes.item(3).text) &".value + 1"
						elseif left(xmldocs.selectSingleNode("//xhead/apclass").text,1)="B" then
							execute "reg.Bap"& year(xi.childNodes.item(3).text) &".value = reg.Bap"& year(xi.childNodes.item(3).text) &".value + 1"
						else
							execute "reg.Cap"& year(xi.childNodes.item(3).text) &".value = reg.Cap"& year(xi.childNodes.item(3).text) &".value + 1"
						end if
					else
						execute "reg.ap"& year(xi.childNodes.item(3).text) &"err.value = reg.ap"& year(xi.childNodes.item(3).text) &"err.value + 1"
						msgremark = msgremark & "("& xi.childNodes.item(1).text &"-"& xi.childNodes.item(2).text
						if trim(xmldocs.selectSingleNode("//xhead/apcust_no").text)<>empty then
							msgremark = msgremark & "無申請人種類("& xmldocs.selectSingleNode("//xhead/apcust_no").text &"))"
						else
							msgremark = msgremark & "無申請人種類("& xmldocs.selectSingleNode("//xhead/apsqlno").text &"))"
						end if
						err = err + 1
					end if
				else
					execute "reg.ap"& year(xi.childNodes.item(3).text) &"err.value = reg.ap"& year(xi.childNodes.item(3).text) &"err.value + 1"
					msgremark = msgremark & "(" & xi.childNodes.item(1).text &"-"& xi.childNodes.item(2).text
					msgremark = msgremark & "無申請人資料)"
					err = err + 1
				end if
			else
				execute "reg.ap"& year(xi.childNodes.item(3).text) &"err.value = reg.ap"& year(xi.childNodes.item(3).text) &"err.value + 1"
				msgremark = msgremark & "(" & xi.childNodes.item(1).text &"-"& xi.childNodes.item(2).text
				msgremark = msgremark & "無申請人資料)"
				err = err + 1
			end if		
			set xmldocs = nothing
		next
		set root = nothing
	end if
	set xmldoc = nothing
	
	reg.remark.value = "(有問題資料共"& err &"筆)" & msgremark
	reg.ap2006sum.value = cdbl(reg.Aap2006.value) + cdbl(reg.Bap2006.value) + cdbl(reg.Cap2006.value)
	reg.ap2007sum.value = cdbl(reg.Aap2007.value) + cdbl(reg.Bap2007.value) + cdbl(reg.Cap2007.value)
	reg.Aapsum.value = cdbl(reg.Aap2006.value) + cdbl(reg.Aap2007.value)
	reg.Bapsum.value = cdbl(reg.Bap2006.value) + cdbl(reg.Bap2007.value)
	reg.Capsum.value = cdbl(reg.Cap2006.value) + cdbl(reg.Cap2007.value)
	reg.apsum.value = cdbl(reg.Aapsum.value) + cdbl(reg.Bapsum.value) + cdbl(reg.Capsum.value)
	reg.aperr.value = cdbl(reg.ap2006err.value) + cdbl(reg.ap2007err.value)
	reg.sum2006.value = cdbl(reg.ap2006sum.value) + cdbl(reg.ap2006err.value)
	reg.sum2007.value = cdbl(reg.ap2007sum.value) + cdbl(reg.ap2007err.value)
	reg.totalsum.value = cdbl(reg.sum2006.value) + cdbl(reg.sum2007.value)
end function
</script>
</body>
</html>   
