<script Language="vbScript">
'D/N單號
function dn_No_onblurx()
	if trim(reg.dn_No.value)=empty then exit function
	'if reg.hdn_flag.value="Y" then
	'	call chkNum1(reg.dn_No,"D/N單號")
	'elseif reg.hdn_flag.value="N" then
	'	call chkNum1(reg.dn_No,"C/N單號")
	'end if
	IF reg.hdn_avg_flag.value<>"Y" then
		call chkdn_no()
	End IF
end function
'D/N日期
function dn_date_onblurx()
	IF trim(reg.dn_date.value) = empty then 
		exit function
	End IF	
	IF chkdateformat(reg.dn_date) = false then
		<%if prgid<>"exp23" and prgid<>"exp25" and prgid<>"exp12" and prgid<>"exp233" then%>
			IF reg.hdn_avg_flag.value<>"Y" then
				call get_dn_rate1
			End IF
			
		<%End IF%>
	End IF
end function
'D/N案件數
function dn_cnt_onblurx()
	IF reg.dn_cnt.value=empty then 
		'msgbox "案件數必須輸入!!!"
		'reg.dn_cnt.focus()
		'exit function
		reg.dn_cnt.value=0
	End IF
	IF cdbl(reg.dn_cnt.value)<=0 then
		msgbox "案件數必須>=1"
		reg.dn_cnt.focus()
		exit function
	End IF
	call chkNum1(reg.dn_cnt,"D/N案件數")
end function

'原幣總金額
function dn_totalmoney_onblurx()
	dn_totalmoney_onblur = false
	IF reg.dn_totalmoney.value=empty then 
		reg.dn_totalmoney.value = 0
	End IF
	call chkNum1(reg.dn_totalmoney,"原幣總金額")
	if reg.hdn_flag.value="Y" then
		if cdbl(reg.dn_totalmoney.value) <= 0 then
			msgbox "原幣總金額必須大於 0"
			reg.dn_totalmoney.focus()
			dn_totalmoney_onblur = true
		Else
			'reg.dn_money.value=reg.dn_totalmoney.value
			'call dn_money_onblur()
		end if
	elseif reg.hdn_flag.value="N" then
		if cdbl(reg.dn_totalmoney.value) >= 0 then
			msgbox "原幣總金額必須小於 0"
			reg.dn_totalmoney.focus()
			dn_totalmoney_onblur = true
		Else
			'reg.dn_money.value=reg.dn_totalmoney.value
			'call dn_money_onblur()
		end if
	end if
end function

'D/N專商合併請款案件數
function dn_ptcnt_onblurx()
	dn_ptcnt_onblur = false
	IF reg.dn_ptcnt.value=empty then 
		'msgbox "專商合併請款案件數必須輸入!!!"
		'reg.dn_ptcnt.focus()
		'exit function
		reg.dn_ptcnt.value=0
	End IF
	IF cdbl(reg.dn_ptcnt.value)<=0 then
		msgbox "專商合併請款案件數必須>=1"
		reg.dn_ptcnt.focus()
		dn_ptcnt_onblur = true
		exit function
	End IF
	call chkNum1(reg.dn_ptcnt,"專商D/N案件數")
end function

'專商原幣總金額
function dn_pttotalmoney_onblurx()
	dn_pttotalmoney_onblur = false
	IF reg.dn_pttotalmoney.value=empty then 
		reg.dn_pttotalmoney.value = 0
	End IF
	call chkNum1(reg.dn_pttotalmoney,"專商原幣總金額")
	if reg.hdn_flag.value="Y" then
		if cdbl(reg.dn_pttotalmoney.value) <= 0 then
			msgbox "專商原幣總金額必須大於 0"
			reg.dn_pttotalmoney.focus()
			dn_pttotalmoney_onblur = true
		end if
	elseif reg.hdn_flag.value="N" then
		if cdbl(reg.dn_pttotalmoney.value) >= 0 then
			msgbox "專商原幣總金額必須小於 0"
			reg.dn_pttotalmoney.focus()
			dn_pttotalmoney_onblur = true
		end if
	end if
end function

'原幣金額
function dn_money_onblurx()
	dn_money_onblur = false
	IF reg.dn_money.value=empty then 
		reg.dn_money.value = 0
	End IF
	call chkNum1(reg.dn_money,"原幣金額")
	if reg.hdn_flag.value="Y" then
		if cdbl(reg.dn_money.value) <= 0 then
			msgbox "原幣金額必須大於 0"
			reg.dn_money.focus()
			dn_money_onblur = true
			exit function
		end if
	elseif reg.hdn_flag.value="N" then
		if cdbl(reg.dn_money.value) >= 0 then
			msgbox "原幣金額必須小於 0"
			reg.dn_money.focus()
			dn_money_onblur = true
			exit function
		end if
	end if
	Call cal_dn_ntmoney
	
end function
'服務費
function dn_service_onblurx()
	dn_service_onblur = false
	IF reg.dn_service.value=empty then
		reg.dn_service.value = 0
	End IF
	IF chkNum1(reg.dn_service,"服務費") = false then
		if reg.hdn_flag.value="Y" then
			if cdbl(reg.dn_service.value) < 0 then
				msgbox "服務費必須大於 0"
				reg.dn_service.focus()
				dn_service_onblur = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if cdbl(reg.dn_service.value) >= 0 then
				msgbox "服務費必須小於 0"
				reg.dn_service.focus()
				dn_service_onblur = true
				exit function
			end if
		end if
		'call cal_dn_money
	End IF
end function
'規費
function dn_fees_onblurx()
	dn_fees_onblur = false
	IF reg.dn_fees.value=empty then 
		reg.dn_fees.value = 0
	End IF
	IF chkNum1(reg.dn_fees,"規費") = false then
		if reg.hdn_flag.value="Y" then
			if cdbl(reg.dn_fees.value) < 0 then
				msgbox "規費必須大於 0"
				reg.dn_fees.focus()
				dn_fees_onblur = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if cdbl(reg.dn_fees.value) > 0 and cdbl(reg.dn_fees.value) <> 0 then
				msgbox "規費必須小於 0"
				reg.dn_fees.focus()
				dn_fees_onblur = true
				exit function
			end if
		end if
		'call cal_dn_money
	End IF
end function
'雜項
function dn_othmoney_onblurx()
	dn_othmoney_onblur = false
	IF reg.dn_othmoney.value = empty then
		reg.dn_othmoney.value = 0
	End IF
	IF chkNum1(reg.dn_othmoney,"雜項") = false then
		if reg.hdn_flag.value="Y" then
			if cdbl(reg.dn_othmoney.value) < 0 then
				msgbox "雜項必須大於 0"
				reg.dn_othmoney.focus()
				dn_othmoney_onblur = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if cdbl(reg.dn_othmoney.value) > 0 and cdbl(reg.dn_fees.value) <> 0 then
				msgbox "雜項必須小於 0"
				reg.dn_othmoney.focus()
				dn_othmoney_onblur = true
				exit function
			end if
		end if
		'call cal_dn_money
	End IF
end function
'郵電費
Function pos_fee_onblurx()
	pos_fee_onblur = false
	IF reg.pos_fee.value = empty then
		reg.pos_fee.value = 0
	End IF
	IF chkNum1(reg.pos_fee,"郵電費") = false then
		if reg.hdn_flag.value="N" then
			if cdbl(reg.pos_fee.value) > 0 then
				msgbox "郵電費必須等於 0"
				reg.pos_fee.focus()
				pos_fee_onblur = true	
				exit function
			end if
		Else
			if cdbl(reg.pos_fee.value) > cdbl(reg.stand_pos_fee.value) then
				msgbox "郵電費不可超過" & reg.stand_pos_fee.value
				reg.pos_fee.focus()
				pos_fee_onblur = true	
				exit function
			end if
		end if
		call cal_dn_nttotal
	End IF
End Function
'手續費
Function hand_fee_onblurx()
	hand_fee_onblur = false
	IF reg.hand_fee.value = empty then
		reg.hand_fee.value = 0
	End IF
	IF chkNum1(reg.hand_fee,"手續費") = false then
		if reg.hdn_flag.value="N" then
			if cdbl(reg.hand_fee.value) > 0 then
				msgbox "手續費必須等於 0"
				reg.hand_fee.focus()
				hand_fee_onblur = true	
				exit function
			end if
		Else
			if cdbl(reg.hand_fee.value) > cdbl(reg.stand_hand_fee.value) then
				msgbox "手續費不可超過" & reg.stand_hand_fee.value
				reg.hand_fee.focus()
				hand_fee_onblur = true	
				exit function
			end if	
		end if
		call cal_dn_nttotal
	End IF
End Function

Function cal_dn_ntmoney()
	IF reg.dn_money.value=empty then
		reg.dn_money.value=0 
	End IF
	<%if prgid<>"exp23" and prgid<>"exp25" then%>
		IF reg.dn_rate.value=empty then
		 reg.dn_rate.value= 0 
		End IF 
		IF reg.dn_money.value <> empty and reg.dn_rate.value <> empty then
			'reg.dn_ntmoney.value = round(cdbl(reg.dn_money.value) * cdbl(reg.dn_rate.value),0)
			reg.dn_ntmoney.value = int((cdbl(reg.dn_money.value) * cdbl(reg.dn_rate.value))+0.5)
		End IF
	<%End IF%>
	call cal_dn_nttotal
End Function

Function cal_dn_nttotal()
	totdn_ntmoney = 0
	IF reg.dn_ntmoney.value <> empty then
		totdn_ntmoney = totdn_ntmoney + reg.dn_ntmoney.value
	End IF
	IF reg.pos_fee.value <> empty then
		totdn_ntmoney = totdn_ntmoney + reg.pos_fee.value
	End IF
	IF reg.hand_fee.value <> empty then
		totdn_ntmoney = totdn_ntmoney + reg.hand_fee.value
	End IF
	reg.dn_nttotal.value = totdn_ntmoney
End Function

Function cal_dn_money()
	totdn_money = 0
	IF reg.dn_service.value <> empty then
		totdn_money = totdn_money + reg.dn_service.value
	End IF
	IF reg.dn_fees.value <> empty then
		totdn_money = totdn_money + reg.dn_fees.value
	End IF
	IF reg.dn_othmoney.value <> empty then
		totdn_money = totdn_money + reg.dn_othmoney.value
	End IF
	reg.dn_money.value = totdn_money
	call cal_dn_ntmoney
End Function


'多筆---------------------------------------------------------------------


'計算總計原幣金額
Function sum_dn_money()
	stot_dn_money=0
	for fi=1 to reg.chknum.value
		IF chkNum1(eval("reg.dn_money"& fi),"原幣金額") = true then
			money_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.dn_money"& fi &".value"))<>"" then
			stot_dn_money = stot_dn_money + cdbl(eval("reg.dn_money"& fi &".value"))
		End IF
	next
	IF money_flag="Y" then
		exit function
	End IF
	reg.tot_tdn_money.value=formatnumber(stot_dn_money,2)
	reg.tot_dn_money.value=stot_dn_money
End Function

'計算總計台幣金額
Function sum_ntdn_money()
	dim i
	
	stot_dn_ntmoney=0
	for i=1 to reg.chknum.value
		IF chkNum1(eval("reg.dn_ntmoney"& i),"台幣金額") = true then
			ntmoney_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.dn_ntmoney"& i &".value"))<>"" then
			stot_dn_ntmoney = stot_dn_ntmoney + cdbl(eval("reg.dn_ntmoney"& i &".value"))
		End IF
	next
	IF ntmoney_flag="Y" then
		exit function
	End IF
	reg.tot_tntdn_money.value = formatnumber(stot_dn_ntmoney,0)
	reg.tot_dn_ntmoney.value = stot_dn_ntmoney
	call sum_tot_tnttotal
End Function

'計算總計服務費
Function sum_dn_service()
	dim i
	
	stot_dn_service=0
	for i=1 to reg.chknum.value
		IF chkNum1(eval("reg.dn_service"& i),"服務費") = true then
			service_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.dn_service"& i &".value"))<>"" then
			stot_dn_service = stot_dn_service + cdbl(eval("reg.dn_service"& i &".value"))
		End IF
	next
	IF service_flag="Y" then
		exit function
	End IF
	reg.tot_tdn_service.value=formatnumber(stot_dn_service,0)
	reg.tot_dn_service.value=stot_dn_service
End Function

'計算總計規費
Function sum_dn_fees()
	dim i
	
	stot_dn_fees=0
	for i=1 to reg.chknum.value
		IF chkNum1(eval("reg.dn_fees"& i),"規費") = true then
			fees_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.dn_fees"& i &".value"))<>"" then
			stot_dn_fees = stot_dn_fees + cdbl(eval("reg.dn_fees"& i &".value"))
		End IF
	next
	IF fees_flag="Y" then
		exit function
	End IF
	reg.tot_tdn_fees.value=formatnumber(stot_dn_fees,0)
	reg.tot_dn_fees.value=stot_dn_fees
End Function

'計算雜費總計金額
Function sum_dn_othmoney()
	dim i
	
	stot_dn_othmoney=0
	for i=1 to reg.chknum.value
		IF chkNum1(eval("reg.dn_othmoney"& i),"雜費") = true then
			othmoney_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.dn_othmoney"& i &".value"))<>"" then
			stot_dn_othmoney = stot_dn_othmoney + cdbl(eval("reg.dn_othmoney"& i &".value"))
		End IF
	next
	IF othmoney_flag="Y" then
		exit function
	End IF
	reg.tot_tdn_othmoney.value=formatnumber(stot_dn_othmoney,0)
	reg.tot_dn_othmoney.value=stot_dn_othmoney
End Function

'抓匯率
Function mget_dn_rate1(pno)
	'tr_yy=year(reg.dn_date.value)
	'tr_mm=month(reg.dn_date.value)
	IF eval("reg.dn_conf_date"&pno&".value")="" then
		tr_yy=year(date())
		tr_mm=month(date())
	Else
		tr_yy=year(eval("reg.dn_conf_date"&pno&".value"))
		tr_mm=month(eval("reg.dn_conf_date"&pno&".value"))
	End IF
	a=eval("reg.dn_currency"&pno&".value")
	SearchSql="Select rate from ex_rate where tr_yy='"& tr_yy &"' and tr_mm='"& tr_mm &"' and currency='"& a &"'"
	url = "../xml/XmlGetSqlData.asp?SearchSql="&SearchSql
	Set xmldoc = CreateObject("Microsoft.XMLDOM")
	xmldoc.async = false
	xmldoc.validateOnParse = true
	If xmldoc.load(url) Then
		Set root = xmldoc.documentElement
		if xmldoc.selectSingleNode("//root/xhead/Found").text="Y" then
			execute "reg.dn_rate"& pno &".value = xmldoc.selectSingleNode(""//root/xhead/rate"").text"
		end if
		set root = nothing
	end if
	set xmldoc = nothing
End Function





'計算單筆的台幣金額
Function mcal_dn_ntmoney(pno)
	IF eval("reg.dn_money"& pno &".value")="" or eval("reg.dn_money"& pno &".value")="0" then exit function
	IF eval("reg.dn_rate"& pno &".value")="" or eval("reg.dn_rate"& pno &".value")="0"  then exit function
	IF eval("reg.dn_money"& pno &".value") <> empty and eval("reg.dn_rate"& pno &".value") <> empty then
		execute "reg.dn_ntmoney"& pno &".value = round(cdbl(reg.dn_money"& pno &".value) * cdbl(reg.dn_rate"& pno &".value),0)"
	End IF
	'台幣金額
	call mcal_dn_nttotal(pno)
	'總計台幣金額
	call sum_ntdn_money()
	'總計原幣金額
	call sum_dn_money
End Function


'總計台幣金額
Function mcal_dn_nttotal(pno)
	mtotdn_ntmoney = 0
	IF eval("reg.dn_ntmoney"& pno &".value") <> empty then
		mtotdn_ntmoney = mtotdn_ntmoney + eval("reg.dn_ntmoney"& pno &".value")
	End IF
	IF eval("reg.pos_fee"& pno &".value") <> empty then
		mtotdn_ntmoney = mtotdn_ntmoney + eval("reg.pos_fee"& pno &".value")
	End IF
	IF eval("reg.hand_fee"& pno &".value") <> empty then
		mtotdn_ntmoney = mtotdn_ntmoney + eval("reg.hand_fee"& pno &".value")
	End IF
	execute "reg.dn_nttotal"& pno &".value = mtotdn_ntmoney "
End Function


'檢查結匯郵電費
Function mcal_pos_fee(pno)
	IF chkNum1(eval("reg.pos_fee"& pno ),"結匯郵電費") = false then
		if reg.hdn_flag.value="Y" then
			if cdbl(eval("reg.pos_fee"& pno &".value"))<=0 then
				msgbox "結匯郵電費必須大於 0"
				execute "reg.pos_fee"& pno &".focus()"
				mcal_pos_fee = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if cdbl(eval("reg.pos_fee"& pno &".value")) > 0 then
				msgbox "結匯郵電費必須等於 0"
				execute "reg.pos_fee"& pno &".focus()"
				mcal_pos_fee = true
				exit function
			end if
		end if
		call sum_tot_tpos_fee
		call mcal_dn_nttotal(pno)
	End IF
End Function

'計算結匯郵電費
Function sum_tot_tpos_fee()
	dim i
	
	stot_tpos_fee=0
	for i=1 to reg.chknum.value
		IF chkNum1(eval("reg.pos_fee"& i),"結匯郵電費") = true then
			tot_tpos_fee_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.pos_fee"& i &".value"))<>"" then
			stot_tpos_fee = stot_tpos_fee + cdbl(eval("reg.pos_fee"& i &".value"))
		End IF
	next
	
	IF tot_tpos_fee_flag="Y" then
		exit function
	End IF
	reg.tot_tpos_fee.value=formatnumber(stot_tpos_fee,0)
	reg.tot_dn_pos_fee.value=stot_tpos_fee
	call sum_tot_tnttotal
End Function

'檢查結匯手續費
Function mcal_hand_fee(pno)
	mcal_hand_fee = false
	IF chkNum1(eval("reg.hand_fee"& pno),"結匯手續費") = false then
		if reg.hdn_flag.value="Y" then
			if eval("reg.hand_fee"& pno &".value")<=0 then
				msgbox "結匯手續費必須大於 0"
				execute "reg.hand_fee"& pno &".focus()"
				mcal_hand_fee = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if eval("reg.hand_fee"& pno &".value") > 0 then
				msgbox "結匯手續費必須等於 0"
				execute "reg.hand_fee"& pno &".focus()"
				mcal_hand_fee = true
				exit function
			end if
		end if
		call sum_tot_thand_fee
		call mcal_dn_nttotal(pno)
	End IF
	
End Function

'計算結匯手續費
Function sum_tot_thand_fee()
	dim i

	stot_tpos_fee=0
	for i=1 to reg.chknum.value
		IF chkNum1(eval("reg.hand_fee"& i),"結匯手續費") = true then
			tot_thand_fee_flag="Y"
			exit for
		End IF
		IF trim(eval("reg.hand_fee"& i &".value"))<>"" then
			stot_thand_fee = stot_thand_fee + cdbl(eval("reg.hand_fee"& i &".value"))
		End IF
	next
	
	IF tot_thand_fee_flag="Y" then
		exit function
	End IF
	reg.tot_thand_fee.value=formatnumber(stot_thand_fee,0)
	reg.tot_dn_hand_fee.value=stot_thand_fee
	call sum_tot_tnttotal
End Function

'服務費
function mdn_service(pno)
	mdn_service = false
	IF chkNum1(eval("reg.dn_service"&pno),"服務費") = false then
		if reg.hdn_flag.value="Y" then
			if eval("reg.dn_service"& pno &".value")<=0 then
				msgbox "服務費必須大於 0"
				execute "reg.dn_service"& pno &".focus()"
				mdn_service = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if eval("reg.dn_service"& pno &".value") > 0 then
				msgbox "服務費必須小於 0"
				execute "reg.dn_service"& pno &".focus()"
				mdn_service = true
				exit function
			end if
		end if
		call sum_cal_dn_money(pno)
	End IF
end function
'規費
function mdn_fees(pno)
	mdn_fees = false
	IF chkNum1(eval("reg.dn_fees"&pno),"規費") = false then
		if reg.hdn_flag.value="Y" then
			if eval("reg.dn_fees"& pno &".value")<=0 then
				msgbox "規費必須大於 0"
				execute "reg.dn_fees"& pno &".focus()"
				mdn_fees = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if eval("reg.dn_fees"& pno &".value") > 0 then
				msgbox "規費必須小於 0"
				execute "reg.dn_fees"& pno &".focus()"
				mdn_fees = true
				exit function
			end if
		end if
		call sum_cal_dn_money(pno)
	End IF
end function
'雜項
function mdn_othmoney(pno)
	mdn_othmoney = false
	IF chkNum1(eval("reg.dn_othmoney"&pno),"雜項") = false then
		if reg.hdn_flag.value="Y" then
			if eval("reg.dn_othmoney"& pno &".value")<=0 then
				msgbox "雜項費用必須大於 0"
				execute "reg.dn_othmoney"& pno &".focus()"
				mdn_othmoney = true
				exit function
			end if
		elseif reg.hdn_flag.value="N" then
			if eval("reg.dn_othmoney"& pno &".value") >= 0 then
				msgbox "雜項費用必須小於 0"
				execute "reg.dn_othmoney"& pno &".focus()"
				mdn_othmoney = true
				exit function
			end if
		end if
		call sum_cal_dn_money(pno)
	End IF
end function

Function sum_cal_dn_money(pno)
	totmdn_money = 0
	IF eval("reg.dn_service"& pno &".value") <> empty then
		execute "totdn_money = totmdn_money + reg.dn_service"& pno &".value"
	End IF
	IF eval("reg.dn_fees"& pno &".value") <> empty then
		execute "totdn_money = totmdn_money + reg.dn_fees"& pno &".value"
	End IF
	IF eval("reg.dn_othmoney"& pno &".value") <> empty then
		execute "totdn_money = totmdn_money + reg.dn_othmoney"& pno &".value"
	End IF
	execute "reg.dn_money"& pno &".value = totmdn_money "
	call mcal_dn_ntmoney
End Function

'服務費+ 規費 +雜項有沒有相等於原幣金額檢查
Function chktdn_money()
	chktdn_money = false
	tdn_money = cdbl(reg.dn_service.value) + cdbl(reg.dn_fees.value) + cdbl(reg.dn_othmoney.value)
	IF cdbl(formatnumber(reg.dn_money.value,2)) <> cdbl(formatnumber(tdn_money,2)) then
		msgbox "原幣金額"&reg.dn_money.value &"＜＞服務費＋規費＋雜項"& tdn_money &"，無法存檔"
		reg.dn_service.focus()
		chktdn_money = true
		exit function
	End IF
End Function	

'計算合計
Function sum_tot_tnttotal()
	stot_tnttotal = 0
	stot_tnttotal = cdbl(reg.tot_dn_ntmoney.value) + cdbl(reg.tot_dn_pos_fee.value) + cdbl(reg.tot_dn_hand_fee.value) 
	reg.tot_tnttotal.value = formatnumber(stot_tnttotal,0)
End Function
'所有項目總金額
Function cal_case_dn_temp_money()
	tempd_dn_case_total_money=0
	for i=1 to reg.tempd_code_num.value
		IF eval("reg.tempd_dn_case_money"& i &".value")="" then
			execute "reg.tempd_dn_case_money"& i &".value=0"
		End IF
		execute "tempd_dn_case_total_money = cdbl(tempd_dn_case_total_money) + cdbl(reg.tempd_dn_case_money"& i &".value)"
	next	
	reg.tempd_dn_case_total_money.value=tempd_dn_case_total_money
End Function
</script>

