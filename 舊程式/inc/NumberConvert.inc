<%
'數字轉成國字
Function NumberToChinese(SendNumber)
   Select Case SendNumber
          Case "1"
               NumberToChinese = "壹"
          Case "2"
               NumberToChinese = "貳"
          Case "3"
               NumberToChinese = "參"
          Case "4"
               NumberToChinese = "肆"
          Case "5"
               NumberToChinese = "伍"
          Case "6"
               NumberToChinese = "陸"
          Case "7"
               NumberToChinese = "柒"
          Case "8"
               NumberToChinese = "捌"
          Case "9"
               NumberToChinese = "玖"
          Case "0"
               NumberToChinese = "零"
   End Select
End Function
'數字轉大寫
Function NumberToCh(SendNumber)
   Select Case SendNumber
          Case "1"
               NumberToCh = "一"
          Case "2"
               NumberToCh = "二"
          Case "3"
               NumberToCh = "三"
          Case "4"
               NumberToCh = "四"
          Case "5"
               NumberToCh = "五"
          Case "6"
               NumberToCh = "六"
          Case "7"
               NumberToCh = "七"
          Case "8"
               NumberToCh = "八"
          Case "9"
               NumberToCh = "九"
          Case "10"
               NumberToCh = "十"
          Case "11"
               NumberToCh = "十一"
          Case "12"
               NumberToCh = "十二"
   End Select
End Function
function getchnum(p1)
	select case len(p1)	
		case 1 '個數
			getchnum = p1
		case 2 '十位數
			if mid(p1,2,1)="零" then
				getchnum = mid(p1,1,1) & "十"
			else
				getchnum = mid(p1,1,1) & "十" & mid(p1,2,1)
			end if
		case 3 '百位數
			if mid(p1,2,1)="零" and mid(p1,3,1)="零" then
				getchnum = mid(p1,1,1) & "百"
			else
				getchnum = mid(p1,1,1) & "百" & mid(p1,2,1) & "十" & mid(p1,3,1)
			end if
		case 4 '千位數
			if mid(p1,2,1)="零" and mid(p1,3,1)="零" and mid(p1,4,1)="零" then
				getchnum = mid(p1,1,1) & "仟"
			else
				getchnum = mid(p1,1,1) & "仟" & mid(p1,2,1) & "百" & mid(p1,3,1) & "十" & mid(p1,4,1)
			end if
	end select
end function
function getchnumToCh(p1)
	select case len(p1)	
		case 1 '個數
			getchnumToCh = NumberToCh(p1)
		case 2 '十位數
			getchnumToCh = NumberToCh(mid(p1,1,1)) & "十" & NumberToCh(mid(p1,2,1))
		case 3 '百位數
			getchnumToCh = NumberToCh(mid(p1,1,1)) & "百" & NumberToCh(mid(p1,2,1)) & "十" & NumberToCh(mid(p1,3,1))
		case 4 '千位數
			getchnumToCh = NumberToCh(mid(p1,1,1)) & "仟" & NumberToCh(mid(p1,2,1)) & "百" & NumberToCh(mid(p1,3,1)) & "十" & NumberToCh(mid(p1,4,1))
	end select
end function
'日期 1~31 中文字
function getChinesenum1(p1)
	dim i
	for i=1 to len(p1)
		getChinesenum1 = getChinesenum1 & NumberToCh(mid(p1,i,1))
	next
	if cint(p1)<10 then
	elseif cint(p1)>=10 and cint(p1)<=19 then
		getChinesenum1 = "十" & mid(getChinesenum1,2,1)
	else
		getChinesenum1 = mid(getChinesenum1,1,1) & "十" & mid(getChinesenum1,2,1)
	end if
end function
'輸入金額轉換成國字
Function gfsTrnCurrencyNumToChineseText(pNumber)
    If Trim(CStr(pNumber)) <> "" Then
        If IsNumeric(pNumber) Then
            nText = Array("零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖", "拾")
            nTextUnit = Array("", "拾", "佰", "仟", "萬", "拾", "佰", "仟", "億", "拾", "佰", "仟", "兆", "拾", "佰")
            If Right(CStr(pNumber), 1) = "." Then
                pNumber = Left(pNumber, Len(pNumber) - 1)
            End If
            If InStr(1, pNumber, ".") = 0 Then
                tLen = Len(CStr(CDbl(pNumber)))
                tNumber = CStr(CDbl(pNumber))
                ReDim nNumArray(tLen)
                For I = 0 To tLen - 1
                    nNumArray(I) = CInt(Right(tNumber, 1))
                    'response.write i&"="&nNumArray(I)&"<br>"
                    tNumber = Left(tNumber, Len(tNumber) - 1)
                Next
                ReDim nChineseNumText(tLen)
                For I = 0 To tLen - 1
                    For j = 0 To UBound(nText) - 1
                        If CInt(nNumArray(I)) = CInt(j) Then
                            nChineseNumText(I) = nText(j)
                        End If
                    Next
                    'response.write i&"="&nChineseNumText(I)&"<br>"
                Next
                For I = 0 To tLen - 1
                    If CInt(nNumArray(I)) = "0" Then
                        If I > 0 Then
                            'if i=4 then
                            '   response.write i&"="&CInt(nNumArray(I))&","&CInt(nNumArray(I - 1))&"<br>"
                            'end if
                            If CInt(nNumArray(I)) = "0" And CInt(nNumArray(I - 1)) = "0" Then
                               if i=4 then	'萬元以上若為0要給單位，不然100736會變成壹拾零柒佰參拾陸元
                                   if CInt(nNumArray(I - 2)) = "0" Then
                                   else
                                      finalText = finalText & nTextUnit(I) 
                                   end if   
                               end if
                            Else
                                if i=4 then'萬元以上若為0要給單位，不然101904會變成壹拾零壹仟玖佰零肆元
                                   finalText = finalText & nTextUnit(I) 
                                      
                                else
                                   finalText = finalText & nChineseNumText(I)
                                end if   
                            End If
                        End If
                    Else
                        finalText = finalText & nTextUnit(I) & nChineseNumText(I)
                    End If
                    'response.write i&"="&finalText&"<br>"
                Next
                finalText = "元" & finalText
                gfsTrnCurrencyNumToChineseText = StrReverse(finalText)
            Else
                gfsTrnCurrencyNumToChineseText = "零元整"
            End If
        Else
            gfsTrnCurrencyNumToChineseText = "零元整"
        End If
    Else
        gfsTrnCurrencyNumToChineseText = "零元整"
    End If
End Function
%>
