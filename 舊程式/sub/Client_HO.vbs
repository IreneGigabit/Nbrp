<script Language="vbScript">
'---H:\_系統檔案\Intranet-brp\系統分析-出專\大陸或英國或歐洲之發明管控香港標準專利案(20080311).ppt
function getHOctrl_date(pnewold,pcase1,pcountry,ppub_date,popen_date,phk_seq,phk_seq1)
	<%if cgrs="AS" then%> exit function <%end if%>
	if pcase1<>"1" then exit function
	'msgbox pcase1
	addmonth = 6 '加6個月
	if pcase1 = "1" and (pcountry = "CM" or pcountry= "EN" or pcountry= "EU") then
		if trim(phk_seq)<>empty and trim(phk_seq)<>"0" then
			ctrl_country = "HO" '管在香港案
		else
			ctrl_country = "EU" '管在大陸或英國或歐洲
		end if
	elseif pcase1 = "1" and pcountry = "HO" then
		if trim(phk_seq)<>empty and trim(phk_seq)<>"0" then
			ctrl_country = "HO" '管在香港案
		end if
	else
		ctrl_country = ""
	end if
	'msgbox ctrl_country
'	if trim(ppub_date) <> empty then
'		if cdate(dateadd("m",addmonth,ppub_date)) > date() then
'			have_ctrl = "Y" '要管制
'			HOctrl_date1 = dateadd("m",addmonth,cdate(ppub_date))
'			if trim(popen_date) <> empty then
'				have_ctrl = "Y" '要管制
'				if dateadd("m",addmonth,popen_date) > date() then
'					HOctrl_date2 = dateadd("m",addmonth,popen_date)
'				end if
'			end if
'		end if
'	else
'		have_ctrl = "N" '不用管制
'	end if
	have_ctrl = "N" '不用管制
	if trim(popen_date) <> empty then '公告日
		have_ctrl = "Y" '要管制
		if dateadd("m",addmonth,popen_date) > date() then
			HOctrl_date2 = dateadd("m",addmonth,popen_date)
		end if
	end if
	if trim(ppub_date) <> empty then '公開日
		if cdate(dateadd("m",addmonth,ppub_date)) > date() then
			have_ctrl = "Y" '要管制
			HOctrl_date1 = dateadd("m",addmonth,cdate(ppub_date))
		end if
	end if
	'msgbox HOctrl_date1 &"-"& HOctrl_date2
	'若已管制過不再重覆管制
	if trim(phk_seq)<>empty and trim(phk_seq1)<>"0" then
		sql = "select ctrl_type,ctrl_date from ctrl_exp"
		sql = sql & " where seq="& phk_seq &" and seq1='"& phk_seq1 &"'"
		sql = sql & " and ctrl_type in ('A9','A10','B9','B10')"
		sql = sql & " union select ctrl_type,ctrl_date from resp_exp"
		sql = sql & " where seq="& phk_seq &" and seq1='"& phk_seq1 &"'"
		sql = sql & " and ctrl_type in ('A9','A10','B9','B10')"
	else
		sql = "select ctrl_type,ctrl_date from ctrl_exp"
		sql = sql & " where seq="& reg.seq.value &" and seq1='"& reg.seq1.value &"'"
		sql = sql & " and ctrl_type in ('A9','A10','B9','B10')"
		sql = sql & " union select ctrl_type,ctrl_date from resp_exp"
		sql = sql & " where seq="& reg.seq.value &" and seq1='"& reg.seq1.value &"'"
		sql = sql & " and ctrl_type in ('A9','A10','B9','B10')"
	end if
	url = "../xml/xmlgetsqldata.asp?searchsql=" & sql
	'window.open url
	set xmldocs = CreateObject("Microsoft.XMLDOM")
	xmldocs.async = false
	xmldocs.validateOnParse = true
	if xmldocs.load (url) then
		if xmldocs.selectSingleNode("//xhead/Found").text = "Y" then
			Set root = xmldocs.documentElement
			For Each xi In root.childNodes
				if xi.childNodes.item(1).text = "A9" or xi.childNodes.item(1).text = "B9" then
					HOctrl_date1 = ""
				end if
				if xi.childNodes.item(1).text = "A10" or xi.childNodes.item(1).text = "B10" then
					HOctrl_date2 = ""
				end if
			next
			set root = nothing
		end if
	end if
	set xmldocs =nothing
	
	'msgbox have_ctrl
	if have_ctrl = "Y" then
		if HOctrl_date1<>empty then
			'msgbox phk_seq
			if (trim(ctrl_country) = "EU" and (trim(phk_seq)=empty or trim(phk_seq)="0")) _
			or trim(pcountry) = "HO" then
				for ixi = reg.ctrlnum.value to 1 step -1
					if eval("reg.date_ctrl"&reg.ctrlnum.value&".value") = "pub_date" then
						tabctrl.deleteRow(ixi+2) '由0起
						reg.ctrlnum.value = reg.ctrlnum.value - 1
					end if
				next
				if HOctrl_date1<>empty then
					ctrl_Add_button_onclick
					execute "reg.ctrl_type"&reg.ctrlnum.value&".value = ""B9"""
					execute "reg.ctrl_date"&reg.ctrlnum.value&".value = HOctrl_date1"
					'execute "reg.ctrl_remark"&reg.ctrlnum.value&".value = ""香港標準專利第一階段申請"""
					execute "reg.date_ctrl"&reg.ctrlnum.value&".value = ""pub_date"""
					execute "reg.sys_flag"&reg.ctrlnum.value&".value= ""※"""
					execute "reg.ctrl_type"&reg.ctrlnum.value&".disabled = true"
					execute "reg.io_flag"&reg.ctrlnum.value&".value = ""Y"""
				end if
			else
				document.all.HOtabctrl.style.display = ""
				document.all.span_HOfseq.innerHtml = "("&"<%=session("se_branch")%>"&"PE-"&reg.hk_seq.value&"-"&reg.hk_seq1.value&")"
				'msgbox reg.HOctrlnum.value
				for ixi = reg.HOctrlnum.value to 1 step -1
					'msgbox ixi+2
					if eval("reg.HOdate_ctrl"&reg.HOctrlnum.value&".value") = "pub_date" then
						HOtabctrl.deleteRow(ixi+1) '由0起
						reg.HOctrlnum.value = reg.HOctrlnum.value - 1
					end if
				next
				if HOctrl_date1<>empty then
					HOctrl_Add_button_onclick
					execute "reg.HOctrl_type"&reg.HOctrlnum.value&".value = ""A9"""
					execute "reg.HOctrl_date"&reg.HOctrlnum.value&".value = HOctrl_date1"
					'execute "reg.HOctrl_remark"&reg.HOctrlnum.value&".value = ""香港標準專利第一階段申請"""
					execute "reg.HOdate_ctrl"&reg.HOctrlnum.value&".value = ""pub_date"""
					execute "reg.HOsys_flag"&reg.HOctrlnum.value&".value= ""※"""
					execute "reg.HOctrl_type"&reg.HOctrlnum.value&".disabled = true"
					execute "reg.HOio_flag"&reg.HOctrlnum.value&".value = ""Y"""
				end if
			end if
		end if

		if HOctrl_date2<>empty then
			'msgbox phk_seq
			if ("<%=cgrs%>"<>"AR" and trim(ctrl_country) = "EU" and (trim(phk_seq)=empty or trim(phk_seq)="0")) _
			or (trim(pcountry) = "HO" and reg.cgrs.value<>"CR") then
				for ixi = reg.ctrlnum.value to 1 step -1
					if eval("reg.date_ctrl"&reg.ctrlnum.value&".value") = "open_date" then
						tabctrl.deleteRow(ixi+2) '由0起
						reg.ctrlnum.value = reg.ctrlnum.value - 1
					end if
				next
				if HOctrl_date2<>empty then
					ctrl_Add_button_onclick
					execute "reg.ctrl_type"&reg.ctrlnum.value&".value = ""B10"""
					execute "reg.ctrl_date"&reg.ctrlnum.value&".value = HOctrl_date2"
					'execute "reg.ctrl_remark"&reg.ctrlnum.value&".value = ""香港標準專利第二階段申請"""
					execute "reg.date_ctrl"&reg.ctrlnum.value&".value = ""open_date"""
					execute "reg.sys_flag"&reg.ctrlnum.value&".value= ""※"""
					execute "reg.ctrl_type"&reg.ctrlnum.value&".disabled = true"
					execute "reg.io_flag"&reg.ctrlnum.value&".value = ""Y"""
				end if
			else
				if (trim(phk_seq)<>empty and trim(phk_seq)<>"0") then
					document.all.HOtabctrl.style.display = ""
					document.all.span_HOfseq.innerHtml = "("&"<%=session("se_branch")%>"&"PE-"&reg.hk_seq.value&"-"&reg.hk_seq1.value&")"
					'msgbox reg.HOctrlnum.value
					for ixi = reg.HOctrlnum.value to 1 step -1
						'msgbox ixi+2
						if eval("reg.HOdate_ctrl"&reg.HOctrlnum.value&".value") = "open_date" then
							HOtabctrl.deleteRow(ixi+1) '由0起
							reg.HOctrlnum.value = reg.HOctrlnum.value - 1
						end if
					next
					if HOctrl_date2<>empty then
						HOctrl_Add_button_onclick
						execute "reg.HOctrl_type"&reg.HOctrlnum.value&".value = ""A10"""
						execute "reg.HOctrl_date"&reg.HOctrlnum.value&".value = HOctrl_date2"
						'execute "reg.HOctrl_remark"&reg.HOctrlnum.value&".value = ""香港標準專利第二階段申請"""
						execute "reg.HOdate_ctrl"&reg.HOctrlnum.value&".value = ""open_date"""
						execute "reg.HOsys_flag"&reg.HOctrlnum.value&".value= ""※"""
						execute "reg.HOctrl_type"&reg.HOctrlnum.value&".disabled = true"
						execute "reg.HOio_flag"&reg.HOctrlnum.value&".value = ""Y"""
					end if
				end if
			end if
		end if
	end if
end function
</script>
