<%@ Page Language="C#" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Collections.Generic"%>
<%@ Import Namespace = "Word=Microsoft.Office.Interop.Word" %>

<script runat="server">
	protected object wdCell=Word.WdUnits.wdCell;
	protected object wdCharacter = Word.WdUnits.wdCharacter;
	protected object wdCharacterFormatting = Word.WdUnits.wdCharacterFormatting;
	protected object wdColumn = Word.WdUnits.wdColumn;
	protected object wdItem = Word.WdUnits.wdItem;
	protected object wdLine = Word.WdUnits.wdLine;
	protected object wdParagraph = Word.WdUnits.wdParagraph;
	protected object wdParagraphFormatting = Word.WdUnits.wdParagraphFormatting;
	protected object wdRow = Word.WdUnits.wdRow;
	protected object wdScreen = Word.WdUnits.wdScreen;
	protected object wdSection = Word.WdUnits.wdSection;
	protected object wdSentence = Word.WdUnits.wdSentence;
	protected object wdStory = Word.WdUnits.wdStory;
	protected object wdTable = Word.WdUnits.wdTable;
	protected object wdWindow = Word.WdUnits.wdWindow;
	protected object wdWord = Word.WdUnits.wdWord;
	protected object wdExtend = 1;

	protected Word._Application wordApp = null;
	protected object oMissing = System.Reflection.Missing.Value;
	protected object oCount1=1;

	private void Page_Load(System.Object sender, System.EventArgs e) {
			Response.Write("$('#chkmsg').html('');\r\n");
			string FileName = Server.MapPath("~/" + Request["catch_path"]);

				if (!File.Exists(FileName)) {
					Response.Write("$('#chkmsg').html('<Font align=left color=\"red\" size=3>找不到說明書Word檔(" + FileName.Replace("\\", "\\\\") + ")!!</font><BR>');\r\n");
					if ((Request["debug"] ?? "").ToUpper() == "Y") {
						Response.Write("$('#chkmsg').append('虛擬目錄:~/" + Request["catch_path"] + "<BR>');\r\n");
						Response.Write("$('#chkmsg').append('轉換後:" + FileName.Replace("\\", "\\\\") + "<BR>');\r\n");
					}
					Response.End();
				}
				wordApp = new Word.Application();

				object oFalse = false;//執行過程不在畫面上開啟 Word
				object oTrue = false;//唯讀模式
				object oFilePath = FileName;    //檔案路徑
				Word._Document myDoc = wordApp.Documents.Open(ref oFilePath, ref oMissing, ref oTrue, ref oMissing,
									ref oMissing, ref oMissing, ref oMissing, ref oMissing,
									ref oMissing, ref oMissing, ref oMissing, ref oFalse,
									ref oMissing, ref oMissing, ref oMissing, ref oMissing);
				myDoc.Activate();
				try {
					Response.Write("var errFlag=false;\r\n");
					
					//20170808 增加檢查案件名稱
					string title_line = Get_name("【");
					title_line = title_line.Replace("【", "").Replace("】", "");
					SQL = " select form_name from cust_code where Code_type='word_tit_p' and code_name='" + title_line + "' ";
					using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
						if (!dr.Read()) {
							Response.Write("$('#chkmsg').append('<Font align=left color=\"red\" size=3>找不到申請書設定，請聯繫資訊人員!!</font><BR>');\r\n");
							if ((Request["debug"] ?? "").ToUpper() == "Y") {
								Response.Write("$('#chkmsg').append('" + SQL.Replace("'", "\\'") + "<BR>');\r\n");
							}
						} else {
							string[] arr_appl = dr.SafeRead("form_name", "").Split('|');//中文專利名稱tag|英文專利名稱tag
							string cappl_line = Get_name(arr_appl[0]);//抓中文專利名稱tag
							string[] split_cappl = cappl_line.Split('】');
							
							//檢查中文專利名稱
							Response.Write("var cappl_name=document.getElementsByName('cappl_name')[0].value;\r\n");
							Response.Write("if (cappl_name.HTMLEncode()!='" + split_cappl[1].Trim() + "'.HTMLEncode()){\r\n");
							Response.Write("	errFlag=true;\r\n");
							Response.Write("	$('#chkmsg').append('<Font align=left color=\"red\" size=3>" + split_cappl[0] + "】申請書案件名稱(" + split_cappl[1].Trim() + ")與案件主檔('+cappl_name+')不符!!</font><BR>');\r\n");
							Response.Write("}\r\n");
						}
					}
					
					//20170808 增加檢查規費
					string fee_line = Get_name("【繳費金額】");
					string[] split_fee = fee_line.Split('】');
					
					if (split_fee.Length == 2) {
						Response.Write("var fee=document.getElementsByName('fees')[0].value;\r\n");
						Response.Write("if (fee!='" + split_fee[1].Trim() + "'){\r\n");
						Response.Write("	errFlag=true;\r\n");
						Response.Write("	$('#chkmsg').append('<Font align=left color=\"red\" size=3>【繳費金額】官發應繳規費('+fee+')與申請書填寫金額(" + split_fee[1].Trim() + ")不符!!</font><BR>');\r\n");
						Response.Write("}\r\n");
					}
					
					//檢查附送書件
					//20170126 原用tagList定義的tag名檢查,改用word【附送書件】區塊,查dmp_attach是否有上傳
					List<string> attachList = Get_AttachBlock();
					for (int z = 0; z < attachList.Count; z++) {
						if (attachList[z] != "") {
							string[] split_line = attachList[z].Replace("　", "").Split('】');
							if (split_line.Length == 2) {
								SQL = " select * from dmp_attach a ";
								SQL += "where seq = '" + Request["seq"] + "' ";
								SQL += " and seq1 = '" + Request["seq1"] + "' ";
								SQL += " and step_grade = '" + Request["step_grade"] + "' ";
								SQL += " and source_name='" + split_line[1].Trim() + "' ";
								SQL += " and esend_flag='Y' ";
								SQL += " and attach_flag<>'D' ";
								using (SqlDataReader dr1 = conn.ExecuteReader(SQL)) {
									if (!dr1.HasRows) {
										Response.Write("errFlag=true;\r\n");
										Response.Write("$('#chkmsg').append('<Font align=left color=\"red\" size=3>" + split_line[0] + "】<b>" + split_line[1] + "</b> 抓取對應附件有錯誤，請檢查附送書件之檔案是否已經上傳 !!</font><BR>');\r\n");
									}
								}
							}
						}
					}

					Response.Write("if (!errFlag){\r\n");
					Response.Write("	$('#chkmsg').html('<Font align=left color=\"darkblue\" size=3>檢查完成，請執行確認!!</font><BR>');\r\n");
					Response.Write("	$('#button0').attr('disabled', true);\r\n");
					Response.Write("}\r\n");
				}
				catch (Exception ex) {
					Response.Write("errFlag=true;\r\n");
					Response.Write("$('#chkmsg').html('<Font align=left color=\"red\" size=3>Eeception - " + ex.Message + "!!</font><BR>');\r\n");
				}
				finally {
					wordApp.ActiveDocument.Close(ref oMissing, ref oMissing, ref oMissing);
					wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
					if (myDoc != null)
						System.Runtime.InteropServices.Marshal.ReleaseComObject(myDoc);
					if (wordApp != null)
						System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
					myDoc = null;
					wordApp = null;
					GC.Collect();
				}
	}

	//抓取摘要(電子申請格式)
	protected string Get_eSummary() {
		string get_value = "";
		wordApp.Selection.HomeKey(ref wdStory, ref oMissing);
		wordApp.Selection.Find.ClearFormatting();
		wordApp.Selection.Find.Text = "【*摘要】";
		wordApp.Selection.Find.Forward = true;
		wordApp.Selection.Find.MatchWholeWord = true;
		wordApp.Selection.Find.MatchWildcards = true;
		
		if (wordApp.Selection.Find.Execute(ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing)) {
			wordApp.Selection.MoveRight(wdCharacter, oCount1);

			wordApp.Selection.Find.Text = "【中文】";
			wordApp.Selection.Find.Forward = true;
			wordApp.Selection.Find.MatchWholeWord = true;
			
			if (wordApp.Selection.Find.Execute(ref oMissing,
			ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
			ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
			ref oMissing, ref oMissing)) {
				wordApp.Selection.MoveRight(wdCharacter, oCount1);

				int i = 0;
				while (++i < 100) {//防止無限迴圈
					wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref oMissing);//ctrl+↓
					wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref wdExtend);//ctrl+shift+↓

					string strTemp = wordApp.Selection.Text;
					strTemp = strTemp.Replace(((char)13).ToString(), "");//整行複製會帶最後的換行符號
					strTemp = strTemp.Replace("　", "");//全形空白
					strTemp = strTemp.Replace(((char)9).ToString(), "");//tab
					strTemp = strTemp.Replace(((char)12).ToString(), "");//換頁
					strTemp = strTemp.Trim();

					if (strTemp.IndexOf("【英文】") > -1 || strTemp.IndexOf("【指定代表圖】")>-1) {
						break;
					} else {
						get_value += strTemp;
					}
				}
			}
		}
		
		return get_value;
	}

	//抓取摘要(紙本申請格式)
	protected string Get_pSummary() {
		string get_value = "";
		wordApp.Selection.HomeKey(ref wdStory, ref oMissing);
		wordApp.Selection.Find.ClearFormatting();
		wordApp.Selection.Find.Text = "中文*摘要：";
		wordApp.Selection.Find.Forward = true;
		wordApp.Selection.Find.MatchWholeWord = true;
		wordApp.Selection.Find.MatchWildcards = true;

		if (wordApp.Selection.Find.Execute(ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing)) {
			wordApp.Selection.MoveRight(wdCharacter, oCount1);

			int i = 0;
			while (++i < 100) {//防止無限迴圈
				wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref oMissing);//ctrl+↓
				wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref wdExtend);//ctrl+shift+↓

				string strTemp = wordApp.Selection.Text;
				strTemp = strTemp.Replace(((char)13).ToString(), "");//整行複製會帶最後的換行符號
				strTemp = strTemp.Replace("　", "");//全形空白
				strTemp = strTemp.Replace(((char)9).ToString(), "");//tab
				strTemp = strTemp.Replace(((char)12).ToString(), "");//換頁
				strTemp = strTemp.Trim();

				if (strTemp.IndexOf("英文") > -1 || strTemp.IndexOf("摘要：") > -1) {
					break;
				} else {
					get_value += strTemp;
				}
			}
		}

		return get_value;
	}

	//抓取專利申請範圍(電子申請格式)
	protected string Get_ERange() {
		string get_value = "";
		wordApp.Selection.HomeKey(ref wdStory, ref oMissing);
		wordApp.Selection.Find.ClearFormatting();
		wordApp.Selection.Find.Text = "【*申請專利範圍】";
		wordApp.Selection.Find.Forward = true;
		wordApp.Selection.Find.MatchWholeWord = true;
		wordApp.Selection.Find.MatchWildcards = true;

		if (wordApp.Selection.Find.Execute(ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing)) {

			int i = 0;
			while (++i < 100) {//防止無限迴圈
				wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref oMissing);//ctrl+↓
				wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref wdExtend);//ctrl+shift+↓

				string strTemp = wordApp.Selection.Text;
				strTemp = strTemp.Replace(((char)13).ToString(), "");//整行複製會帶最後的換行符號
				strTemp = strTemp.Replace("　", "");//全形空白
				strTemp = strTemp.Replace(((char)9).ToString(), "");//tab
				strTemp = strTemp.Replace(((char)12).ToString(), "");//換頁
				strTemp = strTemp.Trim();

				if (wordApp.Selection.Paragraphs[1].Range.ListFormat.ListString != "【第1項】"
					&& (wordApp.Selection.Paragraphs[1].Range.ListFormat.ListString.IndexOf("【") > -1 || strTemp == "")) {
					break;
				} else {
					get_value += strTemp;
				}
			}
		}

		return get_value;
	}

	//抓取專利申請範圍(紙本申請格式)
	protected string Get_PRange() {
		string get_value = "";
		wordApp.Selection.HomeKey(ref wdStory, ref oMissing);
		wordApp.Selection.Find.ClearFormatting();
		wordApp.Selection.Find.Text = "申請專利範圍：";
		wordApp.Selection.Find.Forward = true;
		wordApp.Selection.Find.MatchWholeWord = true;
		wordApp.Selection.Find.MatchWildcards = true;

		if (wordApp.Selection.Find.Execute(ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing)) {

			int i = 0;
			while (++i < 100) {//防止無限迴圈
				wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref oMissing);//ctrl+↓
				wordApp.Selection.MoveDown(ref wdParagraph, ref oCount1, ref wdExtend);//ctrl+shift+↓

				string strTemp = wordApp.Selection.Text;
				strTemp = strTemp.Replace(((char)13).ToString(), "");//整行複製會帶最後的換行符號
				strTemp = strTemp.Replace("　", "");//全形空白
				strTemp = strTemp.Replace(((char)9).ToString(), "");//tab
				strTemp = strTemp.Replace(((char)12).ToString(), "");//換頁
				strTemp = strTemp.Trim();

				if ((wordApp.Selection.Paragraphs[1].Range.ListFormat.ListString != "1" && wordApp.Selection.Paragraphs[1].Range.ListFormat.ListString != "")
					|| (strTemp.IndexOf("2.") > -1 || strTemp == "")) {
					break;
				} else {
					get_value += strTemp;
				}
			}
		}

		return get_value;
	}
</script>
