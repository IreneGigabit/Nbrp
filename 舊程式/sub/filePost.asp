<%
' Adjust this depending on the size 
' of the files you will be accepting
server.scripttimeout = 5400

const ForWriting = 2
const TristateTrue = -1
crlf = chr(13) & chr(10)

' This function retreives a field's name
function getFieldName( infoStr)
	sPos = inStr( infoStr, "name=")
	endPos = inStr( sPos + 6, infoStr, chr(34) & ";")
	if endPos = 0 then
		endPos = inStr( sPos + 6, infoStr, chr(34))
	end if
	getFieldName = mid( infoStr, sPos + 6, endPos - (sPos + 6))
end function

' This function retreives a file field's filename
function getFileName( infoStr)
	sPos = inStr( infoStr, "filename=")
	endPos = inStr( infoStr, chr(34) & crlf)
	getFileName = mid( infoStr, sPos + 10, endPos - (sPos + 10))
end function

' This function retreives a file field's mime type
function getFileType( infoStr)
	sPos = inStr( infoStr, "Content-Type: ")
	getFileType = mid( infoStr, sPos + 14)
end function

' Yank the file (and anything else) that was posted
postData = ""
Dim biData
biData = Request.BinaryRead(Request.TotalBytes)
' Careful! It's binary! So, let's change 
' it into something a bit more manageable
for nIndex = 1 to LenB( biData)
	postData = postData & Chr(AscB(MidB( biData, nIndex, 1)))
next

' Having used BinaryRead, the Request.Form collection is no longer
' available to us. So, we have to parse the request variables ourselves!
' First, let's find that encoding type!
contentType = Request.ServerVariables( "HTTP_CONTENT_TYPE")
ctArray = split( contentType, ";")
' file posts only work well when the encoding is 
' "multipart/form-data", so let's check for that!
if trim(ctArray(0)) = "multipart/form-data" then
	errMsg = ""
	' grab the form boundry...
	bArray = split( trim( ctArray(1)), "=")
	boundry = trim( bArray(1))
	' now use that to split up all the variables!
	formData = split( postData, boundry)
	' now, we need to extract the information for each variable and it's data
	dim myRequest, myRequestFiles(9, 3) 
	Set myRequest = CreateObject("Scripting.Dictionary")
	'Set myRequestFiles = CreateObject("Scripting.Dictionary")
	fileCount = 0
	for x = 0 to ubound( formData)
		' two sets of crlf mark the end of the information about this field
		' everything after that is the value
		infoEnd = instr( formData(x), crlf & crlf)
		if infoEnd > 0 then
			' pull the info for this field, minus the stuff at the ends...
			varInfo = mid( formData(x), 3, infoEnd - 3)
			' pull the value for this field, being sure to 
			' skip the crlf pairs at the start and the crlf-- at the end
			varValue = mid( formData(x), infoEnd + 4, len(formData(x)) - infoEnd - 7)
			' now, is this a file?
			if (instr( varInfo, "filename=") > 0) then
				' place it into our files array
				' While this supports more than one file uploaded at a time
				' we only consider the single file case in this example
				myRequestFiles( fileCount, 0) = getFieldName( varInfo)
				myRequestFiles( fileCount, 1) = varValue
				myRequestFiles( fileCount, 2) = getFileName( varInfo)
				myRequestFiles( fileCount, 3) = getFileType( varInfo)
				fileCount = fileCount + 1
			else
				' it's a regular field
				myRequest.add getFieldName( varInfo), varValue
			end if
		end if
	next
else
	errMsg = "Wrong encoding type!"
end if 

' save the actual posted file
' if you are supporting more than one file, just turn this into a loop!
set lf = server.createObject( "Scripting.FileSystemObject")
if myRequest("filename") = "original" then
	' Use the file name that came with the file
	' at this point, you need to determine what sort of client sent the file
	' Mac's only send the file name, with no path information, while Windows
	' clients send the entire path of the file that was selected
	browserType = UCase( Request.ServerVariables( "HTTP_USER_AGENT"))
	if (inStr(browserType, "WIN") > 0) then
		' it's Windows... yank the filename off the end!
		sPos = inStrRev( myRequestFiles( 0, 2), "\")
		fName = mid( myRequestFiles( 0, 2), sPos + 1)
	end if
	if (inStr(browserType, "MAC") > 0) then
		' it's a Mac. Simple.
		' Of course, Mac file names can contain characters that are 
		' illegal under Windows, so you need to look out for that!
		fName = myRequestFiles(0, 2)
	end if
	' If your upload path is different, set that here
	filePath = "./" & fName
else
	' use the user-specified file name instead
	' If your upload path is different, set that here
	filePath = "./" & myRequest("userSpecifiedName")
end if
savePath = server.mapPath( filePath)
set saveFile = lf.createtextfile(savePath, true)
saveFile.write( myRequestFiles(0, 1))
saveFile.close

' IIS tends to hang if you don't explicitly return SOMETHING.
' So, you should redirect to another page or simply thank them right here...
%>
<html>
<body>
<% if errMsg = "" then %>
	Thanks for the file! It was yummy!
<%else %>
	<%= errMsg %>
<% end if %>
</body>
</html> 
