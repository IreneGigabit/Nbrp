<%@ Language=VBScript CodePage=65001 %>
<%
Session.CodePage = 65001
Response.CharSet = "utf-8"

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set objUpload = New SnoopyUpload '�إߤW�ǹ�H
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
rMrk = "" '�^�ǭȡA�Ǧ^�e���A��#@#���j�A�Ъ`�N���ǧ��� multi_upload_file.asp(uploadSuccess1)���������Ǥ]�n��

seqdept = objUpload.Form("seqdept")  'P���M�BPE�X�M
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

Set objFile = objUpload.File("Filedata")  '�n�Pjs/swfupload.js���]���W�ٲŦX
FName = objFile.FileName  '��l�ɦW
sExt = objFile.FileExt  '���ɦW .pdf
FSize = objFile.FileSize  '�ɮפj�p 



'�]�ǤJ�Ǽ�attach_no���|�֥[�Acreate temp table�Ȧs�A�]�i�קK��H�P�ɾާ@
dim maxattach_no1
dim maxattach_no2
maxattach_no1 = 1
maxattach_no2 = 1
'--1.���R��attachtemp��3�p�ɫe����ơA�A�J�A�H�K�ӿ�W�Ǧ����@�s��attach_no��W
in_date = formatdatetime(DateAdd("h",-3,now),2) &" "& formatdatetime(DateAdd("h",-3,now),4) &":00"
usql = "delete from "& temptable &" where syscode='"& session("syscode") &"' and branch='"& session("se_branch") &"' and dept='"& session("dept") &"'"
usql = usql & " and seq_area='"& seq_area &"' and seq="& seq &" and seq1='"& seq1 &"' and step_grade="& step_grade 
usql = usql & " and in_date<'"& in_date &"'"
conn.Execute usql
'--2.����̤j��
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

'�ɮצW�� AttName=NP-12345--0001-24306-1.doc
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

'sPath=share folder������|
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
'attach_path=����������|
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
    usql = usql & ",'�h�ɤW��')"
end if
rs.Close
conn.Execute usql

'��array�NŪ��Ƽg�^�e���n�`�N����
rMrk = "1#@#" & prgid & "#@#" & seq & "#@#" & seq1 & "#@#" & step_grade & "#@#" & job_sqlno & "#@#" & attach_no
rMrk = rMrk & "#@#" & AttName & "#@#" & sPath & "#@#" & attach_path & "#@#" & FName & "#@#" & cstr(FSize) 
                
Response.Write rMrk
'response.End 
MsgStr = "�W�Ǧ��\ !"
Set objFile = Nothing

Set objUpload = Nothing

Session.CodePage = 950
%>
