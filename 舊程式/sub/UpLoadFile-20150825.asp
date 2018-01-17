<%@ Language=VBScript CodePage=65001 %>
<%
Session.CodePage = 65001
Response.CharSet = "utf-8"

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set objUpload = New SnoopyUpload '建立上傳對象
objUpload.UploadInit "utf-8"

HTProgCode=objUpload.Form("prgid")
HTProgPrefix="UpLoadFile"
HTProgAcs=4

prgid = objUpload.Form("prgid")
%>
<!--#INCLUDE FILE="../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/adovbs.inc" -->
<!--#INCLUDE FILE="../sub/Server_File.asp" -->
<!--#INCLUDE FILE="../sub/SnoopyUpLoad.asp" -->
<%
MsgStr = ""
UpID = ""
FName = ""
FSize = "0"
rMrk = "" '回傳值，傳回畫面，用#@#分隔，請注意順序改變 multi_upload_file.asp(uploadSuccess1)接收的順序也要改

seqdept = objUpload.Form("seqdept")  'P內專、PE出專
seq_area = ""
if seqdept="P" or seqdept="T" then
    seq_area = "I"
elseif seqdept="PE" or seqdept="PE" then
    seq_area = "E"
end if
seq = objUpload.Form("seq")
seq1 = objUpload.Form("seq1")
step_grade = objUpload.Form("step_grade")
job_sqlno = objUpload.Form("job_sqlno")
upfolder = objUpload.Form("upfolder")  '_/123/12345
attach_tablename = objUpload.Form("attach_tablename") 
temptable = objUpload.Form("temptable") 
'attach_no = objUpload.Form("attach_no")
'attach_no = session("attach_no")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open session("btbrtdb")
Set rs = Server.CreateObject("ADODB.Recordset")

Dim objFile
Dim n	
Dim sExt
Dim AttName
Dim fso

call getFileServer(session("se_branch"))
call Check_CreateFolder_virtual(gbrWebDir,upfolder)

Set objFile = objUpload.File("Filedata")  '要與js/swfupload.js中設的名稱符合
FName = objFile.FileName  '原始檔名
sExt = objFile.FileExt  '副檔名 .pdf
FSize = objFile.FileSize  '檔案大小 



'因傳入傳數attach_no不會累加，create temp table暫存，也可避免兩人同時操作
dim maxattach_no1
dim maxattach_no2
maxattach_no1 = 1
maxattach_no2 = 1
'--1.先刪除attachtemp中3小時前的資料，再入，以免承辦上傳但未作存檔attach_no虛增
in_date = formatdatetime(DateAdd("h",-3,now),2) &" "& formatdatetime(DateAdd("h",-3,now),4) &":00"
usql = "delete from "& temptable &" where syscode='"& session("syscode") &"' and branch='"& session("se_branch") &"' and dept='"& session("dept") &"'"
usql = usql & " and seq_area='"& seq_area &"' and seq="& seq &" and seq1='"& seq1 &"' and step_grade="& step_grade 
usql = usql & " and in_date<'"& in_date &"'"
conn.Execute usql
'--2.抓取最大值
isql = "select isnull(max(attach_no),1)+1 as maxattach_no from "& attach_tablename
isql = isql &" where seq_area='"& seq_area &"' and seq="& seq &" and seq1='"& seq1 &"' and step_grade="& step_grade
rs.Open isql, conn, adOpenForwardOnly, adLockReadOnly ,adCmdText
if not rs.EOF then
    maxattach_no1 = rs.Fields(0).Value
end if
rs.Close 
isql = "select isnull(max(attach_no),1)+1 as maxattach_no from "& temptable 
isql = isql &" where syscode='"& session("syscode") &"' and branch='"& session("se_branch") &"' and dept='"& session("dept") &"'"
isql = isql & " and seq_area='"& seq_area &"' and seq="& seq &" and seq1='"& seq1 &"' and step_grade="& step_grade
rs.Open isql, conn, adOpenForwardOnly, adLockReadOnly ,adCmdText
if not rs.EOF then
    maxattach_no2 = rs.Fields(0).Value
end if
rs.Close 
if cdbl(maxattach_no2)>cdbl(maxattach_no1) then
    attach_no = maxattach_no2
else
    attach_no = maxattach_no1
end if

'檔案名稱 AttName=NP-12345--0001-24306-1.doc
AttName = ""	
AttName = session("se_branch")&seqdept&"-"& seq
if seq1="_" then
	AttName = AttName &"-"
else
	AttName = AttName &"-" & seq
end if
AttName = AttName & "-"
AttName = AttName & string(4-len(step_grade),"0")
AttName = AttName & step_grade &"-"&job_sqlno&"-"
AttName = AttName & attach_no &"."& mid(sExt,2)

'sPath=share folder完整路徑
sPath = "\\"& gbrfilesmapervername & "\" & session("se_branch")&seqdept &"\" & upfolder  & "\"
'pfldr1=\\sinn03\NPE  pfldr2=_/174/17432
pfldr1 = "\\"& gbrfilesmapervername &"\"& session("se_branch")&seqdept
pfldr2 = upfolder
'---test begin
'rMrk = "1#@#" & prgid & "#@#" & seq & "#@#" & seq1 & "#@#" & step_grade & "#@#" & job_sqlno & "#@#" & attach_no
'rMrk = rMrk & "#@#" & AttName & "#@#" & sPath & "#@#" & attach_path & "#@#" & FName & "#@#" & cstr(FSize)
'rMrk = rMrk & "#@#" & pfldr1 & "#@#" & pfldr2
'on error resume next
'Response.Write rMrk
'response.End 
'rMrk = err.number &"---"& err.Description 
'---test end
call Check_CreateFolder(pfldr1,pfldr2)
'---test begin
'rMrk = "1#@#" & prgid & "#@#" & seq & "#@#" & seq1 & "#@#" & step_grade & "#@#" & job_sqlno & "#@#" & attach_no
'rMrk = rMrk & "#@#" & AttName & "#@#" & sPath & "#@#" & attach_path & "#@#" & FName & "#@#" & cstr(FSize)
'rMrk = rMrk & "#@#" & pfldr1 & "#@#" & pfldr2
'Response.Write rMrk
'response.End 
'---test end

sPath = sPath & AttName
sPath = replace(sPath,"\","/") 

'AttName:NP-70010--0001-24308-2.pdf
'attach_path:/brp/NP/_/700/70010/NP-70010--0001-24308-3.pdf
'sPath : //web02/NP/_/700/70008/NP-70008--0001-24306-1.pdf
'attach_path=虛擬完整路徑
attach_path = replace(sPath,replace("\\"& gbrfilesmapervername,"\","/"),"/"&session("syscodeiis")) 
objFile.SaveAs sPath

isql = "select attach_no from "& temptable 
isql = isql &" where syscode='"& session("syscode") &"' and branch='"& session("se_branch") &"' and dept='"& session("dept") &"'"
isql = isql & " and seq_area='"& seq_area &"' and seq="& seq &" and seq1='"& seq1 &"' and step_grade="& step_grade
isql = isql & " and in_scode='"& session("scode") &"'"
rs.Open isql, conn, adOpenForwardOnly, adLockReadOnly ,adCmdText
if not rs.EOF then
    usql = "update "& temptable &" set attach_no='"& attach_no &"'"
    usql = usql &" where syscode='"& session("syscode") &"' and branch='"& session("se_branch") &"' and dept='"& session("dept") &"'"
    usql = usql & " and seq_area='"& seq_area &"' and seq="& seq &" and seq1='"& seq1 &"' and step_grade="& step_grade
    usql = usql & " and in_scode='"& session("scode") &"'"
else
    usql = "insert into "& temptable &"(syscode,apcode,branch,dept,seq_area,seq,seq1,step_grade,attach_no,in_date,in_scode,tran_date,tran_scode,remark)"
    usql = usql & " values('"& session("syscode") &"','"& prgid &"','"& session("se_branch") &"','"& session("dept") &"','"& seq_area &"'"
    usql = usql & ","& seq &",'"& seq1 &"',"& step_grade &",'"& attach_no &"',getdate(),'"& session("scode") &"',getdate(),'"& session("scode") &"'"
    usql = usql & ",'多檔上傳')"
end if
rs.Close
conn.Execute usql

'用array將讀資料寫回畫面要注意順序
rMrk = "1#@#" & prgid & "#@#" & seq & "#@#" & seq1 & "#@#" & step_grade & "#@#" & job_sqlno & "#@#" & attach_no
rMrk = rMrk & "#@#" & AttName & "#@#" & sPath & "#@#" & attach_path & "#@#" & FName & "#@#" & cstr(FSize) 
                
Response.Write rMrk
'response.End 
MsgStr = "上傳成功 !"
Set objFile = Nothing

Set objUpload = Nothing

Session.CodePage = 950
%>
