<%
'處理上傳圖檔的部份
Function upin_tech_attach(ptech_br_sqlno,psource)
	dim i
	'目前資料庫中有的最大值
	maxAttach_no = request("uploadfield")&"_maxAttach_no"
	'目前畫面上的最大值
	filenum=request("uploadfield")&"filenum"
	sqlnum=request("uploadfield")&"sqlnum"
	'目前table的筆數
	attach_cnt=request("uploadfield")&"_attach_cnt"
	'欄位名稱
	uploadfield = trim(request("uploadfield"))
	uploadsource = trim(request("uploadsource"))
	
	for i=1 to cint(request(sqlnum))
		select case trim(request(uploadfield & "_dbflag" & i))
			case "A"
				'當上傳路徑不為空的 and attach_sqlno為空的,才需要新增
				IF trim(request(uploadfield&i))<>empty and trim(request(uploadfield & "_attach_sqlno" & i)) ="" then
					fsql = "insert into tech_attach (Tech_br_sqlno,Source" &_
						   ",in_date,in_scode,Attach_no,attach_path,attach_desc" &_
						   ",Attach_name,Attach_size,attach_flag,Mark,tran_date,tran_scode) values (" &_
						   chknumzero(ptech_br_sqlno) &","& _
						   "'"& psource &"','"& date() &"','"& session("scode")&"','"& trim(request(uploadfield & "_attach_no" & i)) &"','"& request(uploadfield&i) &"'," &_
						   "'"& trim(request(uploadfield & "_desc" & i)) &"','"& trim(request(uploadfield & "_name" & i)) &"'," &_
						   "'"& trim(request(uploadfield & "_size" & i)) &"'," &_
						   "'A','',getdate(),'"& session("scode") &"')"
					'response.write "資料庫無資料新增tech_attach"& i &"<br>" &fsql&"<br><br>"
					ConnTech.execute fsql
				end if	
			case "U"
					call insert_log_table(connTech,"U",prgid,"Tech_attach","attach_sqlno",attach_sqlno)
					uSQL = "Update tech_attach set Source='"& psource &"'" &_
					   ",attach_path='"& request(uploadfield&i) &"'" &_
					   ",attach_desc='"& request(uploadfield & "_desc" & i) &"'" &_
					   ",attach_name='"& request(uploadfield & "_name" & i) &"'" &_
					   ",attach_size='"& request(uploadfield & "_size" & i) &"'" &_
					   ",attach_flag='U'" & _
					   ",tran_date=getdate()" &_
					   ",tran_scode='"&  session("scode") &"'" &_
					   " Where attach_sqlno='"& request(uploadfield&"_attach_sqlno"&i) &"'"
					'response.write "更新資料 < Update exp_attach"& i &"=" & uSQL & "<br><br>"
					ConnTech.execute uSQL
			case "D"
				call insert_log_table(connTech,"U",prgid,"Tech_attach","attach_sqlno",attach_sqlno)
				'當attach_sqlno <> empty時,表示db有值,必須刪除data(update attach_flag = 'D')
				if trim(request(uploadfield & "_attach_sqlno" & i)) <> empty then
					dsql = "update tech_attach set attach_flag='D'"
					dsql = dsql & " where attach_sqlno=" & request(uploadfield&"_attach_sqlno"&i)
					'response.write "沒有path刪除該筆資料"& i &"="& dsql &"<br><br>"
					'response.end
					ConnTech.Execute dsql
				else
					'不需要處理,表示原本db就沒有值
				end if
		end select	
	next
	

	'response.end
End Function

'處理上傳圖檔的部份
Function upin_tech_attach_A(ptech_br_sqlno,pbranch,psource)
	dim i
	'目前資料庫中有的最大值
	maxAttach_no = request("uploadfield")&"_maxAttach_no"
	'目前畫面上的最大值
	filenum=request("uploadfield")&"filenum"
	sqlnum=request("uploadfield")&"sqlnum"
	'目前table的筆數
	attach_cnt=request("uploadfield")&"_attach_cnt"
	'欄位名稱
	uploadfield = trim(request("uploadfield"))
	uploadsource = trim(request("uploadsource"))
	
	for i=1 to cint(request(sqlnum))
		pname = split(trim(request(uploadfield&"_name"&i)),".")
		nfilename="TECH-" & trim(pbranch) & "-" & ptech_br_sqlno & "-" & i & "." & pname(1)
		npathall="\brp\Tech_file\" & trim(pbranch) &"/"& ptech_br_sqlno &"\" & nfilename
		npath="\brp\Tech_file\" & trim(pbranch) &"/"& ptech_br_sqlno
		opath="\brp\Tech_file\temp"
		ofilename=trim(request(uploadfield&"_name"&i))
		
		'response.write "舊檔名=" & ofilename & ";舊path=" & opath &"<br>"
		'response.write "新檔名=" & nfilename & ";新path=" & npath &"<br>"
		'當上傳路徑不為空的 and attach_sqlno為空的,才需要新增
		IF trim(request(uploadfield&i))<>empty and trim(request(uploadfield & "_attach_sqlno" & i)) ="" then
			fsql = "insert into tech_attach (Tech_br_sqlno,Source" &_
				   ",in_date,in_scode,Attach_no,attach_path,attach_desc" &_
				   ",Attach_name,Attach_size,attach_flag,Mark,tran_date,tran_scode) values (" &_
				   chknumzero(ptech_br_sqlno) &","& _
				   "'"& psource &"','"& date() &"','"& session("scode")&"','"& trim(request(uploadfield & "_attach_no" & i)) &"','"& npathall &"'," &_
				   "'"& trim(request(uploadfield & "_desc" & i)) &"','"& nfilename &"'," &_
				   "'"& trim(request(uploadfield & "_size" & i)) &"'," &_
				   "'A','',getdate(),'"& session("scode") &"')"
			'response.write "資料庫無資料新增tech_attach"& i &"<br>" &fsql&"<br><br>"
			ConnTech.execute fsql
			Call Check_CreateFolder(npath)
			Call Createfile(opath,npath,ofilename,nfilename)
		end if	
	next
	'response.end
End Function

%>
