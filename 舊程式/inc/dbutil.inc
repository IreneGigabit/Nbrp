<%
function ServerDate(dt)             'SQL Server 端的系統日期，供Compare時使用
   if Len(dt)=0 Or isnull(dt) then
     	ServerDate=""
   else
     	ServerDate = Year(now()) & "/" & Month(now()) & "/" & Day(now())
   end if
end function

function xStdTime(dt)                   '轉成民國年及時間  999/99/99 10:00 給資料型態為DateTime 使用
   if Len(dt)=0 or isnull(dt) then
     	xStdTime=""
   else
   	xyear = right("0"+ cstr((year(dt)-1911)),3)     '補零
        xmonth = right("0"+ cstr(month(dt)),2)
        xday = right("0"+ cstr(day(dt)),2)
        xhour = right("0" + cstr(hour(dt)),2)
        xminute = right("0" + cstr(minute(dt)),2)
        xStdTime = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute
   end if 
end function

function d8date(dt)     '轉成西元年  9999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
		d8date=""
	else
       xy=year(dt)     
       xm=right("0"+ cstr(month(dt)),2)	'補零
       xd=right("0"+ cstr(day(dt)),2)
       d8date=xy & "/" & xm & "/" & xd
   end if 
end function

function d9date(dt)     '轉成西元年informix  9999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
		d9date=""
	else
       xy=year(dt)     
       xm=right("0"+ cstr(month(dt)),2)	'補零
       xd=right("0"+ cstr(day(dt)),2)
       d9date=xm & "/" & xd & "/" & xy
   end if 
end function

function d7date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
		d7date=""
	else
       xy=right("0"+ cstr((year(dt)-1911)),3)     '補零
       xm=right("0"+ cstr(month(dt)),2)
       xd=right("0"+ cstr(day(dt)),2)
       d7date=xy & "/" & xm & "/" & xd
   end if 
end function
function ckSelect(xs1, xs2)
  if xs1=xs2 then
	ckSelect = " selected"
  else
	ckSelect = ""
  end if
end function

FUNCTION CkStr (s, endchar)
  if s="" then
	ckStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	CkStr="'" & s & "'" & endchar
  end if
END FUNCTION

FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="'" & s & "'" & endchar
  end if
END FUNCTION

FUNCTION pkDate (s, endchar)
	if isnull(s) then
		pkDate = "null"  & endchar
	else
		pos = instr(s, " ")
		if pos = 0 then
			pkDate = "'" & s & "'" & endchar
		else
			pkDate = "'" & mid(s, 1, pos-1) & "'" & endchar
		end if
	end if
END FUNCTION

FUNCTION dfs (s)
  if s="" then
	dfs = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	dfs="'" & s & "',"
  end if
END FUNCTION

FUNCTION dfd (s)
  if s="" then
	dfd = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	dfd="'" & s & "',"
  end if
END FUNCTION

FUNCTION dfn (s)
	dfn= s & ","
END FUNCTION

FUNCTION drs (xs)
  s = request(xs)
  if s="" then
	drs = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	drs="'" & s & "',"
  end if
END FUNCTION

FUNCTION drd (xs)
  s = request(xs)
  if s="" then
	drd = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	drd="'" & s & "',"
  end if
END FUNCTION

FUNCTION drn (xs)
s = request(xs)
  if s="" then
	drn = "null,"
  else
	drn= s & ","
  end if
END FUNCTION
%>