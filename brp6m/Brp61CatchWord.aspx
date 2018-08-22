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

		string orgPath = Request["catch_path"];
		if (orgPath.IndexOf(@"/brp/") == 0) {//『/brp/』開頭要換掉
			orgPath = orgPath.Substring(5);
		}

		string FileName = Server.MapPath("~/" + orgPath);
		if (!File.Exists(FileName)) {
			Response.Write("$('#chkmsg').html('<Font align=left color=\"red\" size=3>找不到說明書Word檔(" + FileName.Replace("\\", "\\\\") + ")!!</font><BR>');\r\n");
			if ((Request["debug"] ?? "").ToUpper() == "Y") {
				Response.Write("$('#chkmsg').append('虛擬目錄:~/" + orgPath + "<BR>');\r\n");
				Response.Write("$('#chkmsg').append('轉換後:" + FileName.Replace("\\", "\\\\") + "<BR>');\r\n");
			}
			Response.End();
		}
		wordApp = new Word.Application();

		object oFalse = false;//執行過程不在畫面上開啟 Word
		object oTrue = true;//唯讀模式
		object oFilePath = FileName;    //檔案路徑
		Word._Document myDoc = wordApp.Documents.Open(ref oFilePath, ref oMissing, ref oTrue, ref oMissing,
							ref oMissing, ref oMissing, ref oMissing, ref oMissing,
							ref oMissing, ref oMissing, ref oMissing, ref oFalse,
							ref oMissing, ref oMissing, ref oMissing, ref oMissing);
		myDoc.Activate();
		try {
			Response.Write("var errFlag=false;\r\n");

			//抓取摘要(電子申請格式)
			string summary = Get_eSummary();
			if (summary == "") {
				//抓取摘要(紙本申請格式)
				summary = Get_pSummary();
			}
			Response.Write("document.getElementById('summary_text').innerHTML = '" + summary.Replace("'", "\\'") + "';\r\n");
			if (summary == "") {
				Response.Write("	errFlag=true;\r\n");
				Response.Write("	$('#chkmsg').append('<Font align=left color=\"red\" size=3>找不到摘要!!</font><BR>');\r\n");
			}

			//抓取專利申請範圍(電子申請格式)
			string range = Get_ERange();
			if (range == "") {
				//抓取專利申請範圍(紙本申請格式)
				range = Get_PRange();
			}
			Response.Write("document.getElementById('range_text').innerHTML = '" + range.Replace("'", "\\'") + "';\r\n");
			if (range == "") {
				Response.Write("	errFlag=true;\r\n");
				Response.Write("	$('#chkmsg').append('<Font align=left color=\"red\" size=3>找不到專利申請範圍!!</font><BR>');\r\n");
			}

			Response.Write("if (!errFlag){\r\n");
			Response.Write("	$('#summary_text').focus();\r\n");
			Response.Write("	alert('擷取完成，請確認內容!!');\r\n");
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
			wordApp.Selection.MoveRight(ref wdCharacter, ref oCount1, ref oMissing);

			wordApp.Selection.Find.Text = "【中文】";
			wordApp.Selection.Find.Forward = true;
			wordApp.Selection.Find.MatchWholeWord = true;
			
			if (wordApp.Selection.Find.Execute(ref oMissing,
			ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
			ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
			ref oMissing, ref oMissing)) {
				wordApp.Selection.MoveRight(ref wdCharacter, ref oCount1, ref oMissing);

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
		
		return get_value.ToBig5();
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
			wordApp.Selection.MoveRight(ref wdCharacter, ref oCount1, ref oMissing);

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

		return get_value.ToBig5();
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

		return get_value.ToBig5();
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

		return get_value.ToBig5();
	}
</script>
