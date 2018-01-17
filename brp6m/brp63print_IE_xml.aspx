<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">
	protected string in_scode = "";
	protected string in_no = "";
	protected string branch = "";
	protected string receipt_title = "";

	public IpoReportXml ipoRpt = null;
	public IPOBaseReport baseRpt = null;
	public StringBuilder sbXml = new StringBuilder();
	
	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;
		Response.ContentType = "application/ms-word  charset=utf-8";
		Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5");

		in_scode = (Request["in_scode"] ?? "").ToString();//n100
		in_no = (Request["in_no"] ?? "").ToString();//20170103001
		branch = (Request["branch"] ?? "").ToString();//N
		receipt_title = (Request["receipt_title"] ?? "").ToString();//B

		ipoRpt = new IpoReportXml(Session["btbrtdb"].ToString(), in_scode, in_no, branch, receipt_title);
		baseRpt = new IPOBaseReport(Session["btbrtdb"].ToString(), in_scode, in_no);
		
		WordOut();
		
		ipoRpt.Close();
		baseRpt.Close();
		
		Response.AddHeader("Content-Disposition", "attachment; filename=\"" + ipoRpt.getSeq() + "-發明-" + DateTime.Now.ToString("yyyyMMdd") + ".doc\"");
	}

	protected void WordOut() {
		DataTable dmp = ipoRpt.getDmp();

		//文件開頭
		sbXml.Append(DocHead_1());

		if (dmp.Rows.Count > 0) {
			//標題抬頭
			sbXml.Append(DocBody_1());
			//案由
			string Doc_Body_2 = DocBody_2();
			Doc_Body_2 = Doc_Body_2.Replace("#case_no#", "10000");
			//一併申請實體審查
			if (dmp.Rows[0]["reality"].ToString() == "Y") {
				Doc_Body_2 = Doc_Body_2.Replace("#reality#", "是");
			} else {
				Doc_Body_2 = Doc_Body_2.Replace("#reality#", "否");
			}
			//事務所或申請人案件編號
			Doc_Body_2 = Doc_Body_2.Replace("#seq#", ipoRpt.getSeq() + "-" + dmp.Rows[0]["scode1"].ToString());
			sbXml.Append(Doc_Body_2);
			sbXml.Append(SpaceString());//空白行

			string Doc_Body_3 = DocBody_3();
			//中文發明名稱 / 英文發明名稱
			Doc_Body_3 = Doc_Body_3.Replace("#cappl_name#", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
			Doc_Body_3 = Doc_Body_3.Replace("#eappl_name#", dmp.Rows[0]["eappl_name"].ToString().ToXmlUnicode());
			sbXml.Append(Doc_Body_3);
			sbXml.Append(SpaceString());//空白行

			//產生 申請人
			sbXml.Append(ipoRpt.GetApcustBlock(Dmp_apcust_data()));
			//產生 代理人1 & 代理人2
			sbXml.Append(ipoRpt.GetAgtBlock(Agt_data()));
			//產生 發明人/新型創作/設計人
			sbXml.Append(ipoRpt.GetAntBlock(Ant_data(), "發明人"));

			//主張優惠期
			string Doc_Body_6 = DocBody_6_1();
			string exh_date = "";
			if (dmp.Rows[0]["exhibitor"].ToString() == "Y") {//參展或發表日期填入表中的發生日期
				if (dmp.Rows[0]["exh_date"] != System.DBNull.Value && dmp.Rows[0]["exh_date"] != null) {
					exh_date = Convert.ToDateTime(dmp.Rows[0]["exh_date"]).ToString("yyyy/MM/dd");
				}
			}
			sbXml.Append(Doc_Body_6.Replace("#exh_date#", exh_date));
			sbXml.Append(SpaceString());//空白行

			//產生 主張優先權 迴圈
			sbXml.Append(ipoRpt.GetPriorBlock(DocBody_7(), DocBody_7_1(), DocBody_7_2(), SpaceString()));

			//主張利用生物材料
			string Doc_Body_7_3 = DocBody_7_3();
			sbXml.Append(DocBody_7_3());
			sbXml.Append(SpaceString());//空白行

			//生物材料不須寄存
			sbXml.Append(DocBody_8());
			sbXml.Append(SpaceString());//空白行

			//聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
			string Doc_Body_81 = DocBody_81();
			if (dmp.Rows[0]["same_apply"].ToString() == "Y") {
				Doc_Body_81 = Doc_Body_81.Replace("#same_apply#", "是");
			} else {
				Doc_Body_81 = Doc_Body_81.Replace("#same_apply#", "");
			}
			sbXml.Append(Doc_Body_81);
			sbXml.Append(SpaceString());//空白行

			//中文本資訊 ,外文本資訊 ,繳費資訊
			sbXml.Append(DocBody_9());
			//20170524 增加收據抬頭選項
			sbXml.Append(Dmp_receipt_title().Replace("#rectitle_name#", ipoRpt.getRectitleName(receipt_title)));
			sbXml.Append(SpaceString());//空白行
			//附送書件
			sbXml.Append(DocBody_10().Replace("#seq#", ipoRpt.getSeq()));
			sbXml.Append(SpaceString());//空白行

			//sw.Append(DocFooter());
			sbXml.Append(DocNewPageFooter());
			//sw.Append(DocTail_1());

			//基本資料表
			baseRpt.Build("發明人");
			sbXml.Append(baseRpt.sbXml);
		}
		//文件結尾
		sbXml.Append(DocTail());
	}
	

		//空白行
		private string SpaceString() {
			return
			"<w:p wsp:rsidR=\"008515F6\" " +
		"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
		"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr></w:r></w:p>";
		}

		private string DocHead_1() {
			return
					"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
			"<?mso-application progid=\"Word.Document\"?>" +
			"<w:wordDocument xmlns:aml=\"http://schemas.microsoft.com/aml/2001/core\" xmlns:dt=\"uuid:C2F41010-65B3-11d1-A29F-00AA00C14882\" xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
			"xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
			"xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\" xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\" " +
			"xmlns:wsp=\"http://schemas.microsoft.com/office/word/2003/wordml/sp2\" xmlns:sl=\"http://schemas.microsoft.com/schemaLibrary/2003/core\" " +
			"w:macrosPresent=\"no\" w:embeddedObjPresent=\"no\" w:ocxPresent=\"no\" xml:space=\"preserve\"><w:ignoreSubtree w:val=\"http://schemas.microsoft.com/office/word/2003/wordml/sp2\"/>" +
			"<w:fonts><w:defaultFonts w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times " +
			"New Roman\" w:cs=\"Times New Roman\"/><w:font w:name=\"Times New Roman\"><w:panose-1 w:val=\"02020603050405020304\"/><w:charset w:val=\"00\"/>" +
			"<w:family w:val=\"Roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb-0=\"20002A87\" w:usb-1=\"80000000\" w:usb-2=\"00000008\" w:usb-3=\"00000000\" " +
			"w:csb-0=\"000001FF\" w:csb-1=\"00000000\"/></w:font><w:font w:name=\"新細明體\"><w:altName w:val=\"PMingLiU\"/><w:panose-1 w:val=\"02020300000000000000\"/>" +
			"<w:charset w:val=\"88\"/><w:family w:val=\"Roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb-0=\"00000003\" w:usb-1=\"080E0000\" w:usb-2=\"00000016\" " +
			"w:usb-3=\"00000000\" w:csb-0=\"00100001\" w:csb-1=\"00000000\"/></w:font><w:font w:name=\"Cambria Math\"><w:panose-1 w:val=\"02040503050406030204\"/>" +
			"<w:charset w:val=\"00\"/><w:family w:val=\"Roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb-0=\"E00002FF\" w:usb-1=\"420024FF\" w:usb-2=\"00000000\" " +
			"w:usb-3=\"00000000\" w:csb-0=\"0000019F\" w:csb-1=\"00000000\"/></w:font><w:font w:name=\"Cambria\"><w:panose-1 w:val=\"02040503050406030204\"/>" +
			"<w:charset w:val=\"00\"/><w:family w:val=\"Roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb-0=\"E00002FF\" w:usb-1=\"400004FF\" w:usb-2=\"00000000\" " +
			"w:usb-3=\"00000000\" w:csb-0=\"0000019F\" w:csb-1=\"00000000\"/></w:font><w:font w:name=\"Calibri\"><w:panose-1 w:val=\"020F0502020204030204\"/>" +
			"<w:charset w:val=\"00\"/><w:family w:val=\"Swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb-0=\"E10002FF\" w:usb-1=\"4000ACFF\" w:usb-2=\"00000009\" " +
			"w:usb-3=\"00000000\" w:csb-0=\"0000019F\" w:csb-1=\"00000000\"/></w:font><w:font w:name=\"@新細明體\"><w:panose-1 w:val=\"02020300000000000000\"/>" +
			"<w:charset w:val=\"88\"/><w:family w:val=\"Roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb-0=\"00000003\" w:usb-1=\"080E0000\" w:usb-2=\"00000016\" " +
			"w:usb-3=\"00000000\" w:csb-0=\"00100001\" w:csb-1=\"00000000\"/></w:font></w:fonts><w:lists>" +
			"<w:listDef w:listDefId=\"0\"><w:lsid w:val=\"05ED05A0\"/>" +
			"<w:plt w:val=\"HybridMultilevel\"/><w:tmpl w:val=\"675C8C10\"/><w:lvl w:ilvl=\"0\" w:tplc=\"75104D78\"><w:start w:val=\"1\"/><w:lvlText " +
			"w:val=\"【主張優先權】\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"480\" w:hanging=\"480\"/></w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times New Roman\" w:hint=\"fareast\"/><w:b w:val=\"off\"/><w:i w:val=\"off\"/>" +
			"<w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%2、\"/>" +
			"<w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"960\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"2\" w:tplc=\"0409001B\"><w:start " +
			"w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"3\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText w:val=\"%4.\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
			"<w:ind w:left=\"1920\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"4\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/>" +
			"<w:lvlText w:val=\"%5、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2400\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"5\" " +
			"w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%6.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"2880\" " +
			"w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"6\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc " +
			"w:val=\"left\"/><w:pPr><w:ind w:left=\"3360\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"7\" w:tplc=\"04090019\"><w:start " +
			"w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%8、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"3840\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"8\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%9.\"/><w:lvlJc " +
			"w:val=\"right\"/><w:pPr><w:ind w:left=\"4320\" w:hanging=\"480\"/></w:pPr></w:lvl></w:listDef>" +
			"<w:listDef w:listDefId=\"1\"><w:lsid " +
			"w:val=\"0E0A38EF\"/><w:plt w:val=\"HybridMultilevel\"/><w:tmpl w:val=\"2A02EAF2\"/><w:lvl w:ilvl=\"0\" w:tplc=\"DD3253EE\"><w:start " +
			"w:val=\"1\"/><w:lvlText w:val=\"【代理人】\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"480\" w:hanging=\"480\"/></w:pPr>" +
			"<w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times New Roman\" w:hint=\"fareast\"/><w:b w:val=\"off\"/>" +
			"<w:i w:val=\"off\"/><w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/>" +
			"<w:lvlText w:val=\"%2、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"960\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"2\" " +
			"w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"1440\" " +
			"w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"3\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText w:val=\"%4.\"/><w:lvlJc " +
			"w:val=\"left\"/><w:pPr><w:ind w:left=\"1920\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"4\" w:tplc=\"04090019\"><w:start " +
			"w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%5、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2400\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"5\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%6.\"/><w:lvlJc " +
			"w:val=\"right\"/><w:pPr><w:ind w:left=\"2880\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"6\" w:tplc=\"0409000F\"><w:start " +
			"w:val=\"1\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"3360\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl " +
			"w:ilvl=\"7\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%8、\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
			"<w:ind w:left=\"3840\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"8\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/>" +
			"<w:lvlText w:val=\"%9.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"4320\" w:hanging=\"480\"/></w:pPr></w:lvl></w:listDef>" +

			"<w:listDef w:listDefId=\"2\"><w:lsid w:val=\"249D468C\"/><w:plt w:val=\"HybridMultilevel\"/><w:tmpl w:val=\"C270F68A\"/><w:lvl w:ilvl=\"0\" w:tplc=\"07000BB0\">" +
			"<w:start w:val=\"1\"/><w:lvlText w:val=\"【主張利用生物材料%1】\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"480\" " +
			"w:hanging=\"480\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times New Roman\" w:hint=\"fareast\"/>" +
			"<w:b w:val=\"off\"/><w:i w:val=\"off\"/><w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\" w:tplc=\"04090019\"><w:start w:val=\"1\"/>" +
			"<w:nfc w:val=\"30\"/><w:lvlText w:val=\"%2、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"960\" w:hanging=\"480\"/></w:pPr></w:lvl>" +
			"<w:lvl w:ilvl=\"2\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"right\"/>" +
			"<w:pPr><w:ind w:left=\"1440\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"3\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText " +
			"w:val=\"%4.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1920\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"4\" w:tplc=\"04090019\">" +
			"<w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%5、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2400\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"5\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%6.\"/><w:lvlJc " +
			"w:val=\"right\"/><w:pPr><w:ind w:left=\"2880\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"6\" w:tplc=\"0409000F\"><w:start " +
			"w:val=\"1\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"3360\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl " +
			"w:ilvl=\"7\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%8、\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
			"<w:ind w:left=\"3840\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"8\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/>" +
			"<w:lvlText w:val=\"%9.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"4320\" w:hanging=\"480\"/></w:pPr></w:lvl></w:listDef>" +

			"<w:listDef w:listDefId=\"3\"><w:lsid w:val=\"29A844E5\"/><w:plt w:val=\"HybridMultilevel\"/><w:tmpl w:val=\"128837EA\"/><w:lvl w:ilvl=\"0\" w:tplc=\"0608C004\">" +
			"<w:start w:val=\"1\"/><w:lvlText w:val=\"【主張優惠期%1】\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"480\" w:hanging=\"480\"/>" +
			"</w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times New Roman\" w:hint=\"fareast\"/><w:b " +
			"w:val=\"off\"/><w:i w:val=\"off\"/><w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\" w:tplc=\"04090019\"><w:start w:val=\"1\"/>" +
			"<w:nfc w:val=\"30\"/><w:lvlText w:val=\"%2、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"960\" w:hanging=\"480\"/></w:pPr></w:lvl>" +
			"<w:lvl w:ilvl=\"2\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"right\"/>" +
			"<w:pPr><w:ind w:left=\"1440\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"3\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText " +
			"w:val=\"%4.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1920\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"4\" w:tplc=\"04090019\">" +
			"<w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%5、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2400\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"5\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%6.\"/><w:lvlJc " +
			"w:val=\"right\"/><w:pPr><w:ind w:left=\"2880\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"6\" w:tplc=\"0409000F\"><w:start " +
			"w:val=\"1\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"3360\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl " +
			"w:ilvl=\"7\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%8、\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
			"<w:ind w:left=\"3840\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"8\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/>" +
			"<w:lvlText w:val=\"%9.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"4320\" w:hanging=\"480\"/></w:pPr></w:lvl></w:listDef>" +

			"<w:listDef w:listDefId=\"4\"><w:lsid w:val=\"3EAE5824\"/><w:plt w:val=\"HybridMultilevel\"/><w:tmpl w:val=\"ADC040D6\"/><w:lvl w:ilvl=\"0\" w:tplc=\"C4801CB4\">" +
			"<w:start w:val=\"1\"/><w:lvlText w:val=\"【申請人】\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"480\" w:hanging=\"480\"/>" +
			"</w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times New Roman\" w:hint=\"fareast\"/><w:b " +
			"w:val=\"off\"/><w:i w:val=\"off\"/><w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\" w:tplc=\"04090019\"><w:start w:val=\"1\"/>" +
			"<w:nfc w:val=\"30\"/><w:lvlText w:val=\"%2、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"960\" w:hanging=\"480\"/></w:pPr></w:lvl>" +
			"<w:lvl w:ilvl=\"2\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"right\"/>" +
			"<w:pPr><w:ind w:left=\"1440\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"3\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText " +
			"w:val=\"%4.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1920\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"4\" w:tplc=\"04090019\">" +
			"<w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%5、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2400\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"5\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%6.\"/><w:lvlJc " +
			"w:val=\"right\"/><w:pPr><w:ind w:left=\"2880\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"6\" w:tplc=\"0409000F\"><w:start " +
			"w:val=\"1\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"3360\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl " +
			"w:ilvl=\"7\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%8、\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
			"<w:ind w:left=\"3840\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"8\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/>" +
			"<w:lvlText w:val=\"%9.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"4320\" w:hanging=\"480\"/></w:pPr></w:lvl></w:listDef>" +
			"<w:listDef w:listDefId=\"5\"><w:lsid w:val=\"7B0A0C09\"/><w:plt w:val=\"HybridMultilevel\"/><w:tmpl w:val=\"D1903E14\"/><w:lvl w:ilvl=\"0\" w:tplc=\"A5948AC2\">" +
			"<w:start w:val=\"1\"/><w:lvlText w:val=\"【發明人】\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"480\" w:hanging=\"480\"/>" +
			"</w:pPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:fareast=\"新細明體\" w:h-ansi=\"Times New Roman\" w:hint=\"fareast\"/><w:b " +
			"w:val=\"off\"/><w:i w:val=\"off\"/><w:sz w:val=\"24\"/></w:rPr></w:lvl><w:lvl w:ilvl=\"1\" w:tplc=\"04090019\"><w:start w:val=\"1\"/>" +
			"<w:nfc w:val=\"30\"/><w:lvlText w:val=\"%2、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"960\" w:hanging=\"480\"/></w:pPr></w:lvl>" +
			"<w:lvl w:ilvl=\"2\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%3.\"/><w:lvlJc w:val=\"right\"/>" +
			"<w:pPr><w:ind w:left=\"1440\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"3\" w:tplc=\"0409000F\"><w:start w:val=\"1\"/><w:lvlText " +
			"w:val=\"%4.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1920\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"4\" w:tplc=\"04090019\">" +
			"<w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%5、\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2400\" w:hanging=\"480\"/>" +
			"</w:pPr></w:lvl><w:lvl w:ilvl=\"5\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/><w:lvlText w:val=\"%6.\"/><w:lvlJc " +
			"w:val=\"right\"/><w:pPr><w:ind w:left=\"2880\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"6\" w:tplc=\"0409000F\"><w:start " +
			"w:val=\"1\"/><w:lvlText w:val=\"%7.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"3360\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl " +
			"w:ilvl=\"7\" w:tplc=\"04090019\"><w:start w:val=\"1\"/><w:nfc w:val=\"30\"/><w:lvlText w:val=\"%8、\"/><w:lvlJc w:val=\"left\"/><w:pPr>" +
			"<w:ind w:left=\"3840\" w:hanging=\"480\"/></w:pPr></w:lvl><w:lvl w:ilvl=\"8\" w:tplc=\"0409001B\"><w:start w:val=\"1\"/><w:nfc w:val=\"2\"/>" +
			"<w:lvlText w:val=\"%9.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"4320\" w:hanging=\"480\"/></w:pPr></w:lvl></w:listDef><w:list " +
			"w:ilfo=\"1\"><w:ilst w:val=\"4\"/></w:list><w:list w:ilfo=\"2\"><w:ilst w:val=\"4\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride " +
			"w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"1\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"2\">" +
			"<w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"3\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride " +
			"w:ilvl=\"4\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"5\"><w:startOverride w:val=\"1\"/></w:lvlOverride>" +
			"<w:lvlOverride w:ilvl=\"6\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"7\"><w:startOverride w:val=\"1\"/>" +
			"</w:lvlOverride><w:lvlOverride w:ilvl=\"8\"><w:startOverride w:val=\"1\"/></w:lvlOverride></w:list><w:list w:ilfo=\"3\"><w:ilst w:val=\"1\"/>" +
			"</w:list><w:list w:ilfo=\"4\"><w:ilst w:val=\"1\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride " +
			"w:ilvl=\"1\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"2\"><w:startOverride w:val=\"1\"/></w:lvlOverride>" +
			"<w:lvlOverride w:ilvl=\"3\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"4\"><w:startOverride w:val=\"1\"/>" +
			"</w:lvlOverride><w:lvlOverride w:ilvl=\"5\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"6\"><w:startOverride " +
			"w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"7\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"8\">" +
			"<w:startOverride w:val=\"1\"/></w:lvlOverride></w:list><w:list w:ilfo=\"5\"><w:ilst w:val=\"5\"/></w:list><w:list w:ilfo=\"6\"><w:ilst " +
			"w:val=\"5\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"1\"><w:startOverride " +
			"w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"2\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"3\">" +
			"<w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"4\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride " +
			"w:ilvl=\"5\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"6\"><w:startOverride w:val=\"1\"/></w:lvlOverride>" +
			"<w:lvlOverride w:ilvl=\"7\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"8\"><w:startOverride w:val=\"1\"/>" +
			"</w:lvlOverride></w:list><w:list w:ilfo=\"7\"><w:ilst w:val=\"3\"/></w:list><w:list w:ilfo=\"8\"><w:ilst w:val=\"3\"/><w:lvlOverride " +
			"w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"1\"><w:startOverride w:val=\"1\"/></w:lvlOverride>" +
			"<w:lvlOverride w:ilvl=\"2\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"3\"><w:startOverride w:val=\"1\"/>" +
			"</w:lvlOverride><w:lvlOverride w:ilvl=\"4\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"5\"><w:startOverride " +
			"w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"6\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"7\">" +
			"<w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"8\"><w:startOverride w:val=\"1\"/></w:lvlOverride></w:list><w:list " +
			"w:ilfo=\"9\"><w:ilst w:val=\"0\"/></w:list><w:list w:ilfo=\"10\"><w:ilst w:val=\"0\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride " +
			"w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"1\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"2\">" +
			"<w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"3\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride " +
			"w:ilvl=\"4\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"5\"><w:startOverride w:val=\"1\"/></w:lvlOverride>" +
			"<w:lvlOverride w:ilvl=\"6\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"7\"><w:startOverride w:val=\"1\"/>" +
			"</w:lvlOverride><w:lvlOverride w:ilvl=\"8\"><w:startOverride w:val=\"1\"/></w:lvlOverride></w:list><w:list w:ilfo=\"11\"><w:ilst w:val=\"2\"/>" +
			"</w:list><w:list w:ilfo=\"12\"><w:ilst w:val=\"2\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride " +
			"w:ilvl=\"1\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"2\"><w:startOverride w:val=\"1\"/></w:lvlOverride>" +
			"<w:lvlOverride w:ilvl=\"3\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"4\"><w:startOverride w:val=\"1\"/>" +
			"</w:lvlOverride><w:lvlOverride w:ilvl=\"5\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"6\"><w:startOverride " +
			"w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"7\"><w:startOverride w:val=\"1\"/></w:lvlOverride><w:lvlOverride w:ilvl=\"8\">" +
			"<w:startOverride w:val=\"1\"/></w:lvlOverride></w:list></w:lists><w:styles><w:versionOfBuiltInStylenames w:val=\"7\"/><w:latentStyles " +
			"w:defLockedState=\"off\" w:latentStyleCount=\"267\"><w:lsdException w:name=\"Normal\"/><w:lsdException w:name=\"heading 1\"/><w:lsdException " +
			"w:name=\"heading 2\"/><w:lsdException w:name=\"heading 3\"/><w:lsdException w:name=\"heading 4\"/><w:lsdException w:name=\"heading " +
			"5\"/><w:lsdException w:name=\"heading 6\"/><w:lsdException w:name=\"heading 7\"/><w:lsdException w:name=\"heading 8\"/><w:lsdException " +
			"w:name=\"heading 9\"/><w:lsdException w:name=\"toc 1\"/><w:lsdException w:name=\"toc 2\"/><w:lsdException w:name=\"toc 3\"/><w:lsdException " +
			"w:name=\"toc 4\"/><w:lsdException w:name=\"toc 5\"/><w:lsdException w:name=\"toc 6\"/><w:lsdException w:name=\"toc 7\"/><w:lsdException " +
			"w:name=\"toc 8\"/><w:lsdException w:name=\"toc 9\"/><w:lsdException w:name=\"caption\"/><w:lsdException w:name=\"Title\"/><w:lsdException " +
			"w:name=\"Default Paragraph Font\"/><w:lsdException w:name=\"Subtitle\"/><w:lsdException w:name=\"Hyperlink\"/><w:lsdException w:name=\"Strong\"/>" +
			"<w:lsdException w:name=\"Emphasis\"/><w:lsdException w:name=\"Table Grid\"/><w:lsdException w:name=\"Placeholder Text\"/><w:lsdException " +
			"w:name=\"No Spacing\"/><w:lsdException w:name=\"Light Shading\"/><w:lsdException w:name=\"Light List\"/><w:lsdException w:name=\"Light " +
			"Grid\"/><w:lsdException w:name=\"Medium Shading 1\"/><w:lsdException w:name=\"Medium Shading 2\"/><w:lsdException w:name=\"Medium " +
			"List 1\"/><w:lsdException w:name=\"Medium List 2\"/><w:lsdException w:name=\"Medium Grid 1\"/><w:lsdException w:name=\"Medium Grid " +
			"2\"/><w:lsdException w:name=\"Medium Grid 3\"/><w:lsdException w:name=\"Dark List\"/><w:lsdException w:name=\"Colorful Shading\"/>" +
			"<w:lsdException w:name=\"Colorful List\"/><w:lsdException w:name=\"Colorful Grid\"/><w:lsdException w:name=\"Light Shading Accent 1\"/>" +
			"<w:lsdException w:name=\"Light List Accent 1\"/><w:lsdException w:name=\"Light Grid Accent 1\"/><w:lsdException w:name=\"Medium Shading " +
			"1 Accent 1\"/><w:lsdException w:name=\"Medium Shading 2 Accent 1\"/><w:lsdException w:name=\"Medium List 1 Accent 1\"/><w:lsdException " +
			"w:name=\"Revision\"/><w:lsdException w:name=\"List Paragraph\"/><w:lsdException w:name=\"Quote\"/><w:lsdException w:name=\"Intense " +
			"Quote\"/><w:lsdException w:name=\"Medium List 2 Accent 1\"/><w:lsdException w:name=\"Medium Grid 1 Accent 1\"/><w:lsdException w:name=\"Medium " +
			"Grid 2 Accent 1\"/><w:lsdException w:name=\"Medium Grid 3 Accent 1\"/><w:lsdException w:name=\"Dark List Accent 1\"/><w:lsdException " +
			"w:name=\"Colorful Shading Accent 1\"/><w:lsdException w:name=\"Colorful List Accent 1\"/><w:lsdException w:name=\"Colorful Grid Accent " +
			"1\"/><w:lsdException w:name=\"Light Shading Accent 2\"/><w:lsdException w:name=\"Light List Accent 2\"/><w:lsdException w:name=\"Light " +
			"Grid Accent 2\"/><w:lsdException w:name=\"Medium Shading 1 Accent 2\"/><w:lsdException w:name=\"Medium Shading 2 Accent 2\"/><w:lsdException " +
			"w:name=\"Medium List 1 Accent 2\"/><w:lsdException w:name=\"Medium List 2 Accent 2\"/><w:lsdException w:name=\"Medium Grid 1 Accent " +
			"2\"/><w:lsdException w:name=\"Medium Grid 2 Accent 2\"/><w:lsdException w:name=\"Medium Grid 3 Accent 2\"/><w:lsdException w:name=\"Dark " +
			"List Accent 2\"/><w:lsdException w:name=\"Colorful Shading Accent 2\"/><w:lsdException w:name=\"Colorful List Accent 2\"/><w:lsdException " +
			"w:name=\"Colorful Grid Accent 2\"/><w:lsdException w:name=\"Light Shading Accent 3\"/><w:lsdException w:name=\"Light List Accent 3\"/>" +
			"<w:lsdException w:name=\"Light Grid Accent 3\"/><w:lsdException w:name=\"Medium Shading 1 Accent 3\"/><w:lsdException w:name=\"Medium " +
			"Shading 2 Accent 3\"/><w:lsdException w:name=\"Medium List 1 Accent 3\"/><w:lsdException w:name=\"Medium List 2 Accent 3\"/><w:lsdException " +
			"w:name=\"Medium Grid 1 Accent 3\"/><w:lsdException w:name=\"Medium Grid 2 Accent 3\"/><w:lsdException w:name=\"Medium Grid 3 Accent " +
			"3\"/><w:lsdException w:name=\"Dark List Accent 3\"/><w:lsdException w:name=\"Colorful Shading Accent 3\"/><w:lsdException w:name=\"Colorful " +
			"List Accent 3\"/><w:lsdException w:name=\"Colorful Grid Accent 3\"/><w:lsdException w:name=\"Light Shading Accent 4\"/><w:lsdException " +
			"w:name=\"Light List Accent 4\"/><w:lsdException w:name=\"Light Grid Accent 4\"/><w:lsdException w:name=\"Medium Shading 1 Accent 4\"/>" +
			"<w:lsdException w:name=\"Medium Shading 2 Accent 4\"/><w:lsdException w:name=\"Medium List 1 Accent 4\"/><w:lsdException w:name=\"Medium " +
			"List 2 Accent 4\"/><w:lsdException w:name=\"Medium Grid 1 Accent 4\"/><w:lsdException w:name=\"Medium Grid 2 Accent 4\"/><w:lsdException " +
			"w:name=\"Medium Grid 3 Accent 4\"/><w:lsdException w:name=\"Dark List Accent 4\"/><w:lsdException w:name=\"Colorful Shading Accent " +
			"4\"/><w:lsdException w:name=\"Colorful List Accent 4\"/><w:lsdException w:name=\"Colorful Grid Accent 4\"/><w:lsdException w:name=\"Light " +
			"Shading Accent 5\"/><w:lsdException w:name=\"Light List Accent 5\"/><w:lsdException w:name=\"Light Grid Accent 5\"/><w:lsdException " +
			"w:name=\"Medium Shading 1 Accent 5\"/><w:lsdException w:name=\"Medium Shading 2 Accent 5\"/><w:lsdException w:name=\"Medium List 1 " +
			"Accent 5\"/><w:lsdException w:name=\"Medium List 2 Accent 5\"/><w:lsdException w:name=\"Medium Grid 1 Accent 5\"/><w:lsdException " +
			"w:name=\"Medium Grid 2 Accent 5\"/><w:lsdException w:name=\"Medium Grid 3 Accent 5\"/><w:lsdException w:name=\"Dark List Accent 5\"/>" +
			"<w:lsdException w:name=\"Colorful Shading Accent 5\"/><w:lsdException w:name=\"Colorful List Accent 5\"/><w:lsdException w:name=\"Colorful " +
			"Grid Accent 5\"/><w:lsdException w:name=\"Light Shading Accent 6\"/><w:lsdException w:name=\"Light List Accent 6\"/><w:lsdException " +
			"w:name=\"Light Grid Accent 6\"/><w:lsdException w:name=\"Medium Shading 1 Accent 6\"/><w:lsdException w:name=\"Medium Shading 2 Accent " +
			"6\"/><w:lsdException w:name=\"Medium List 1 Accent 6\"/><w:lsdException w:name=\"Medium List 2 Accent 6\"/><w:lsdException w:name=\"Medium " +
			"Grid 1 Accent 6\"/><w:lsdException w:name=\"Medium Grid 2 Accent 6\"/><w:lsdException w:name=\"Medium Grid 3 Accent 6\"/><w:lsdException " +
			"w:name=\"Dark List Accent 6\"/><w:lsdException w:name=\"Colorful Shading Accent 6\"/><w:lsdException w:name=\"Colorful List Accent " +
			"6\"/><w:lsdException w:name=\"Colorful Grid Accent 6\"/><w:lsdException w:name=\"Subtle Emphasis\"/><w:lsdException w:name=\"Intense " +
			"Emphasis\"/><w:lsdException w:name=\"Subtle Reference\"/><w:lsdException w:name=\"Intense Reference\"/><w:lsdException w:name=\"Book " +
			"Title\"/><w:lsdException w:name=\"Bibliography\"/><w:lsdException w:name=\"TOC Heading\"/></w:latentStyles><w:style w:type=\"paragraph\" " +
			"w:default=\"on\" w:styleId=\"a\"><w:name w:val=\"Normal\"/><wx:uiName wx:val=\"內文\"/><w:rPr><w:rFonts w:ascii=\"Calibri\" w:h-ansi=\"Calibri\" " +
			"w:cs=\"新細明體\"/><wx:font wx:val=\"Calibri\"/><w:sz w:val=\"24\"/><w:sz-cs w:val=\"24\"/><w:lang w:val=\"EN-US\" w:fareast=\"ZH-TW\" " +
			"w:bidi=\"AR-SA\"/></w:rPr></w:style><w:style w:type=\"character\" w:default=\"on\" w:styleId=\"a0\"><w:name w:val=\"Default Paragraph " +
			"Font\"/><wx:uiName wx:val=\"預設段落字型\"/></w:style><w:style w:type=\"table\" w:default=\"on\" w:styleId=\"a1\"><w:name w:val=\"Normal " +
			"Table\"/><wx:uiName wx:val=\"表格內文\"/><w:rPr><wx:font wx:val=\"Times New Roman\"/><w:lang w:val=\"EN-US\" w:fareast=\"ZH-TW\" " +
			"w:bidi=\"AR-SA\"/></w:rPr><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" " +
			"w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style><w:style " +
			"w:type=\"list\" w:default=\"on\" w:styleId=\"a2\"><w:name w:val=\"No List\"/><wx:uiName wx:val=\"無清單\"/></w:style><w:style w:type=\"character\" " +
			"w:styleId=\"a3\"><w:name w:val=\"Hyperlink\"/><wx:uiName wx:val=\"超連結\"/><w:rPr><w:color w:val=\"0000FF\"/><w:u w:val=\"single\"/>" +
			"</w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"a4\"><w:name w:val=\"FollowedHyperlink\"/><wx:uiName wx:val=\"已查閱的超連結\"/>" +
			"<w:rPr><w:color w:val=\"800080\"/><w:u w:val=\"single\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"a5\"><w:name w:val=\"header\"/>" +
			"<wx:uiName wx:val=\"頁首\"/><w:basedOn w:val=\"a\"/><w:link w:val=\"a6\"/><w:pPr><w:snapToGrid w:val=\"off\"/></w:pPr><w:rPr><wx:font " +
			"wx:val=\"Calibri\"/><w:sz w:val=\"20\"/><w:sz-cs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"a6\"><w:name " +
			"w:val=\"頁首 字元\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"a5\"/><w:locked/></w:style><w:style w:type=\"paragraph\" w:styleId=\"a7\">" +
			"<w:name w:val=\"footer\"/><wx:uiName wx:val=\"頁尾\"/><w:basedOn w:val=\"a\"/><w:link w:val=\"a8\"/><w:pPr><w:snapToGrid w:val=\"off\"/>" +
			"</w:pPr><w:rPr><wx:font wx:val=\"Calibri\"/><w:sz w:val=\"20\"/><w:sz-cs w:val=\"20\"/></w:rPr></w:style><w:style w:type=\"character\" " +
			"w:styleId=\"a8\"><w:name w:val=\"頁尾 字元\"/><w:basedOn w:val=\"a0\"/><w:link w:val=\"a7\"/><w:locked/></w:style><w:style w:type=\"paragraph\" " +
			"w:styleId=\"a9\"><w:name w:val=\"Balloon Text\"/><wx:uiName wx:val=\"註解方塊文字\"/><w:basedOn w:val=\"a\"/><w:link w:val=\"aa\"/>" +
			"<w:rPr><w:rFonts w:ascii=\"Cambria\" w:h-ansi=\"Cambria\" w:cs=\"Times New Roman\"/><wx:font wx:val=\"Cambria\"/><w:sz w:val=\"20\"/>" +
			"<w:sz-cs w:val=\"20\"/><w:lang/></w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"aa\"><w:name w:val=\"註解方塊文字 " +
			"字元\"/><w:link w:val=\"a9\"/><w:locked/><w:rPr><w:rFonts w:ascii=\"Cambria\" w:h-ansi=\"Cambria\" w:hint=\"default\"/></w:rPr></w:style>" +
			"<w:style w:type=\"paragraph\" w:styleId=\"ab\"><w:name w:val=\"Revision\"/><wx:uiName wx:val=\"修訂\"/><w:rPr><w:rFonts w:ascii=\"Calibri\" " +
			"w:h-ansi=\"Calibri\" w:cs=\"新細明體\"/><wx:font wx:val=\"Calibri\"/><w:sz w:val=\"24\"/><w:sz-cs w:val=\"24\"/><w:lang w:val=\"EN-US\" " +
			"w:fareast=\"ZH-TW\" w:bidi=\"AR-SA\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"ac\"><w:name w:val=\"List Paragraph\"/>" +
			"<wx:uiName wx:val=\"清單段落\"/><w:basedOn w:val=\"a\"/><w:pPr><w:ind w:left=\"480\"/></w:pPr><w:rPr><wx:font wx:val=\"Calibri\"/>" +
			"</w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"msochpdefault\"><w:name w:val=\"msochpdefault\"/><w:basedOn w:val=\"a\"/>" +
			"<w:pPr><w:spacing w:before=\"100\" w:before-autospacing=\"on\" w:after=\"100\" w:after-autospacing=\"on\"/></w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/><w:sz w:val=\"20\"/><w:sz-cs w:val=\"20\"/></w:rPr>" +
			"</w:style></w:styles><w:shapeDefaults><o:shapedefaults v:ext=\"edit\" spidmax=\"3074\"/><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" " +
			"data=\"1\"/></o:shapelayout></w:shapeDefaults><w:docPr><w:view w:val=\"print\"/><w:zoom w:percent=\"90\"/><w:doNotEmbedSystemFonts/>" +
			"<w:bordersDontSurroundHeader/><w:bordersDontSurroundFooter/><w:proofState w:grammar=\"clean\"/><w:defaultTabStop w:val=\"480\"/><w:characterSpacingControl " +
			"w:val=\"CompressPunctuation\"/><w:optimizeForBrowser/><w:targetScreenSz w:val=\"1024x768\"/><w:validateAgainstSchema/><w:saveInvalidXML " +
			"w:val=\"off\"/><w:ignoreMixedContent w:val=\"off\"/><w:alwaysShowPlaceholderText w:val=\"off\"/><w:hdrShapeDefaults><o:shapedefaults " +
			"v:ext=\"edit\" spidmax=\"3074\"/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:type=\"separator\"><w:p wsp:rsidR=\"00BF474D\" wsp:rsidRDefault=\"00BF474D\">" +
			"<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=\"continuation-separator\"><w:p wsp:rsidR=\"00BF474D\" wsp:rsidRDefault=\"00BF474D\">" +
			"<w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotePr><w:endnotePr><w:endnote w:type=\"separator\"><w:p wsp:rsidR=\"00BF474D\" " +
			"wsp:rsidRDefault=\"00BF474D\"><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type=\"continuation-separator\"><w:p wsp:rsidR=\"00BF474D\" " +
			"wsp:rsidRDefault=\"00BF474D\"><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotePr><w:compat><w:breakWrappedTables/>" +
			"<w:useFELayout/></w:compat><wsp:rsids><wsp:rsidRoot wsp:val=\"000940B1\"/><wsp:rsid wsp:val=\"000940B1\"/><wsp:rsid wsp:val=\"00097D83\"/>" +
			"<wsp:rsid wsp:val=\"000A07BC\"/><wsp:rsid wsp:val=\"001548CA\"/><wsp:rsid wsp:val=\"00173215\"/><wsp:rsid wsp:val=\"0017701A\"/><wsp:rsid " +
			"wsp:val=\"001B4EDE\"/><wsp:rsid wsp:val=\"001D5812\"/><wsp:rsid wsp:val=\"002563D4\"/><wsp:rsid wsp:val=\"002831B1\"/><wsp:rsid wsp:val=\"00290EC6\"/>" +
			"<wsp:rsid wsp:val=\"002E6DC4\"/><wsp:rsid wsp:val=\"003067C7\"/><wsp:rsid wsp:val=\"00325C57\"/><wsp:rsid wsp:val=\"003304A1\"/><wsp:rsid " +
			"wsp:val=\"0037287A\"/><wsp:rsid wsp:val=\"00397393\"/><wsp:rsid wsp:val=\"00397406\"/><wsp:rsid wsp:val=\"003D0D41\"/><wsp:rsid wsp:val=\"004542AC\"/>" +
			"<wsp:rsid wsp:val=\"004C1689\"/><wsp:rsid wsp:val=\"005029F5\"/><wsp:rsid wsp:val=\"00517763\"/><wsp:rsid wsp:val=\"00521150\"/><wsp:rsid " +
			"wsp:val=\"00534587\"/><wsp:rsid wsp:val=\"005476FC\"/><wsp:rsid wsp:val=\"005477FE\"/><wsp:rsid wsp:val=\"0057572F\"/><wsp:rsid wsp:val=\"007B726F\"/>" +
			"<wsp:rsid wsp:val=\"007F3823\"/><wsp:rsid wsp:val=\"00802186\"/><wsp:rsid wsp:val=\"008406A1\"/><wsp:rsid wsp:val=\"00840A0E\"/><wsp:rsid " +
			"wsp:val=\"00881A76\"/><wsp:rsid wsp:val=\"00890066\"/><wsp:rsid wsp:val=\"008E1167\"/><wsp:rsid wsp:val=\"008E4F69\"/><wsp:rsid wsp:val=\"00975CB3\"/>" +
			"<wsp:rsid wsp:val=\"009E422D\"/><wsp:rsid wsp:val=\"00A53646\"/><wsp:rsid wsp:val=\"00A55D33\"/><wsp:rsid wsp:val=\"00A73EC2\"/><wsp:rsid " +
			"wsp:val=\"00A90BF0\"/><wsp:rsid wsp:val=\"00AA0311\"/><wsp:rsid wsp:val=\"00B01C92\"/><wsp:rsid wsp:val=\"00B13187\"/><wsp:rsid wsp:val=\"00BA7D2C\"/>" +
			"<wsp:rsid wsp:val=\"00BB3E24\"/><wsp:rsid wsp:val=\"00BB5802\"/><wsp:rsid wsp:val=\"00BE731D\"/><wsp:rsid wsp:val=\"00BF474D\"/><wsp:rsid " +
			"wsp:val=\"00C1673F\"/><wsp:rsid wsp:val=\"00C41BA3\"/><wsp:rsid wsp:val=\"00C95ED8\"/><wsp:rsid wsp:val=\"00D8143C\"/><wsp:rsid wsp:val=\"00DA4A7C\"/>" +
			"<wsp:rsid wsp:val=\"00E30BC0\"/><wsp:rsid wsp:val=\"00E3365A\"/><wsp:rsid wsp:val=\"00E76F19\"/><wsp:rsid wsp:val=\"00E83593\"/><wsp:rsid " +
			"wsp:val=\"00E86378\"/><wsp:rsid wsp:val=\"00E91440\"/><wsp:rsid wsp:val=\"00EA7C1C\"/><wsp:rsid wsp:val=\"00ED4953\"/><wsp:rsid wsp:val=\"00EF3DB7\"/>" +
			"<wsp:rsid wsp:val=\"00EF5BA9\"/><wsp:rsid wsp:val=\"00F37F28\"/><wsp:rsid wsp:val=\"00FF7CD5\"/></wsp:rsids></w:docPr><w:body>";
		}

		//標題抬頭
		private string DocBody_1() {
			return
		"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:jc w:val=\"center\"/><w:rPr>" +
		"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/><w:sz w:val=\"44\"/><w:sz-cs w:val=\"44\"/>" +
		"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
		"<w:sz w:val=\"44\"/><w:sz-cs w:val=\"44\"/></w:rPr><w:t>【發明專利申請書】</w:t></w:r></w:p><w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
		"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
		"</w:rPr></w:pPr></w:p>";
		}

		//案由,一併申請實體審查,事務所或申請人案件編號
		private string DocBody_2() {
			return
				"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRPr=\"00565151\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"00565151\"><w:pPr>" +
			"<w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【案由】　　　　　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/></w:rPr>" +
			"<w:t>#case_no#</w:t></w:r></w:p><w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【一併申請實體審查】　　　　　</w:t>" +
			"</w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr><w:t>#reality#</w:t></w:r>" +
			"</w:p><w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【事務所或申請人案件編號】　　</w:t></w:r><w:r wsp:rsidR=\"00565151\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#</w:t>" +
			"</w:r></w:p>";
		}

		//中文發明名稱,英文發明名稱
		private string DocBody_3() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【中文發明名稱】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#cappl_name#</w:t></w:r></w:p><w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【英文發明名稱】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#eappl_name#</w:t>" +
			"</w:r></w:p>";
		}

		//申請人區塊
		private string Dmp_apcust_data() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"00565151\" " +
			"wsp:rsidP=\"00565151\"><w:pPr><w:pStyle w:val=\"ac\"/><w:ind w:left=\"0\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【申請人#apply_num#】</w:t></w:r></w:p><w:p " +
			"wsp:rsidR=\"00565151\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"00B55A67\"><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"8028\"/></w:tabs>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr>" +
			"<w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【國籍】　　　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#ap_country#</w:t></w:r></w:p><w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【#ap_cname1_title#】　　　　　　　</w:t>" +
			"</w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr><w:t>#ap_cname1#</w:t></w:r></w:p><w:p wsp:rsidR=\"00565151\" " +
			"wsp:rsidRDefault=\"00565151\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【#ap_ename1_title#】　　　　　　　#ap_ename1#</w:t></w:r></w:p>" + SpaceString();
		}

		//代理人
		private string Agt_data() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"00565151\" wsp:rsidP=\"00565151\"><w:pPr><w:pStyle w:val=\"ac\"/><w:ind w:left=\"0\"/>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【代理人#agt_num#】</w:t>" +
			"</w:r></w:p><w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r>" +
			"<w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【中文姓名】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#agt_name#</w:t></w:r></w:p>" + SpaceString();
		}

		//發明人迴圈
		private string Ant_data() {
			return "<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"00565151\" wsp:rsidP=\"00565151\"><w:pPr><w:pStyle w:val=\"ac\"/><w:ind w:left=\"0\"/>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【#ant_num#】</w:t></w:r></w:p><w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【國籍】　　　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#ant_country#</w:t></w:r></w:p><w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【中文姓名】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00565151\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#ant_cname#</w:t></w:r></w:p><w:p wsp:rsidR=\"00565151\" wsp:rsidRPr=\"00565151\" " +
			"wsp:rsidRDefault=\"00565151\" wsp:rsidP=\"00565151\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【英文姓名】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#ant_ename#</w:t></w:r></w:p>" + SpaceString();
		}

		//主張優惠期
		private string DocBody_6() {
			return
				"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"00C667B1\"><w:pPr><w:pStyle w:val=\"ac\"/><w:listPr><w:ilvl w:val=\"0\"/><w:ilfo w:val=\"8\"/>" +
			"<wx:t wx:val=\"【主張優惠期1】\"/><wx:font wx:val=\"新細明體\"/></w:listPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE " +
			"w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【發生日期】　　　　　　　#exh_date#</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【因實驗而公開者】　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【因於刊物發表者】　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"005B309B\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【因陳列於政府主辦或認可之展覽會者】</w:t></w:r></w:p>";
		}

		//優惠期事實 20170504 智慧局取消主張優惠期改用優惠期事實
		private string DocBody_6_1() {
			return
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【本案符合優惠期相關規定】　　是／否</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【優惠期事實1】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【發生日期】　　　　　　　#exh_date#</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" " +
			"wsp:rsidRDefault=\"005B309B\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【公開事由】　　　　　　　</w:t></w:r></w:p>";
		}

		//主張優先權迴圈
		private string DocBody_7() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"00CC1819\" wsp:rsidP=\"00CC1819\"><w:pPr><w:pStyle w:val=\"ac\"/><w:ind w:left=\"0\"/>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【主張優先權#prior_num#】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r>" +
			"<w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【申請日】　　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00CC1819\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#prior_date#</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【受理國家或地區】　　　　</w:t></w:r><w:r wsp:rsidR=\"00A67C95\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#prior_country#</w:t>" +
			"</w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【申請案號】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00A67C95\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#prior_no#</w:t>" +
			"</w:r></w:p>";
		}

		//主張優先權迴圈-JA
		private string DocBody_7_1() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:overflowPunct w:val=\"off\"/>" +
			"<w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【專利類別】　　　　　　　</w:t>" +
			"</w:r><w:r wsp:rsidR=\"00A67C95\"><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr><w:t>#case1nm#</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【存取碼】　　　　　　　　</w:t>" +
			"</w:r><w:r wsp:rsidR=\"00A67C95\"><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr><w:t>#mprior_access#</w:t></w:r></w:p>";
		}

		//主張優先權迴圈-KO
		private string DocBody_7_2() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【存取碼】　　　　　　　　</w:t>" +
			"</w:r><w:r wsp:rsidR=\"00A67C95\"><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr><w:t>#mprior_access#</w:t></w:r></w:p>";
		}

		//主張利用生物材料
		private string DocBody_7_3() {
			return
			"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"00C667B1\">" +
			"<w:pPr><w:pStyle w:val=\"ac\"/><w:listPr><w:ilvl w:val=\"0\"/><w:ilfo w:val=\"12\"/><wx:t wx:val=\"【主張利用生物材料1】\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:listPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/>" +
			"<w:spacing w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"005B309B\"><w:pPr>" +
			"<w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【寄存國家】　　　　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"005B309B\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【寄存機構】　　　　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"005B309B\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【寄存日期】　　　　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"00C667B1\" wsp:rsidRDefault=\"005B309B\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【寄存號碼】　　　　　　　</w:t></w:r></w:p>";
		}

		//生物材料不須寄存
		private string DocBody_8() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【生物材料不須寄存】　　　所屬技術領域中具有通常知識者易於獲得。</w:t>" +
			"</w:r></w:p>";
		}

		//聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
		private string DocBody_81() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【聲明本人就相同創作在申請本發明專利之同日-另申請新型專利】　#same_apply#</w:t>" +
			"</w:r></w:p>";
		}

		//中文本資訊 外文本資訊 繳費資訊
		private string DocBody_9() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【中文本資訊】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【摘要頁數】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t>" +
			"</w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:overflowPunct w:val=\"off\"/>" +
			"<w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【說明書頁數】　　　　　　0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN " +
			"w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr>" +
			"<w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【申請專利範圍頁數】　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN " +
			"w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr>" +
			"<w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【圖式頁數】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN " +
			"w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr>" +
			"<w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【頁數總計】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\"><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN " +
			"w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr>" +
			"<w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【申請專利範圍項數】　　　0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【圖式圖數】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t>" +
			"</w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【附英文摘要】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t></w:t>" +
			"</w:r></w:p>" +

			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr></w:p>" +

			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【外文本資訊】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r>" +
			"<w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【外文頁數總計】　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【外文本種類】　　　　　　日文／英文／德文／韓文／法文／俄文／葡萄牙文／西班牙文／阿拉伯文</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t></w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【簡體字本資訊】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r>" +
			"<w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【簡體字頁數總計】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t></w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【繳費資訊】</w:t></w:r>" +
			"</w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【繳費金額】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>0</w:t>" +
			"</w:r></w:p>";
		}

		//20170524 增加收據抬頭選項
		private string Dmp_receipt_title() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【收據抬頭】　　　　　　　</w:t></w:r><w:r wsp:rsidR=\"00B55A67\">" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#rectitle_name#</w:t>" +
			"</w:r></w:p>";
		}

		//附送書件&備註
		private string DocBody_10() {
			return
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【備註】　　　　　　　　　　　</w:t></w:r></w:p>" +
			SpaceString() +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【附送書件】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【基本資料表】　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>#seq#-Contact.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr>" +
			"<w:t>　　【發明摘要】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>#seq#-desc_Abstract.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"<w:t>　　【發明說明書】　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-desc_Description.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【序列表】　　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-Squence.pdf</w:t></w:r></w:p>" +

			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【發明申請專利範圍】　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-desc_Claims.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【發明圖式】　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-Drawings.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>　　【外文本】　　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-ForeignAbstract.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-ForeignDescription.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" " +
			"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-ForeignClaims.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-ForeignDrawings.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-ForeignSpec.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-SimplifiedAbstract.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-SimplifiedDescription.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-SimplifiedClaims.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-SimplifiedDrawings.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-SimplifiedSpec.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【國際優先權證明文件】　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-Priority.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【優惠期證明文件】　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-ICExperiment.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【優惠期證明文件】　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-Exhibition.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times " +
			"New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【委任書】　　　　　　　　</w:t>" +
			"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>#seq#-PowerAttorney.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【國內生物材料寄存證明文件】</w:t></w:r><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-FIRDI99999.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【國外生物材料寄存證明文件】</w:t></w:r><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-ATCC99999.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
			"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【生物材料為通常知識者易於獲得證明文件】</w:t></w:r><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-EasilyObtained.pdf</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"003D09F2\" wsp:rsidRDefault=\"003D09F2\" wsp:rsidP=\"003D09F2\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
			"w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
			"<wx:font wx:val=\"新細明體\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　【其他】</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"003D09F2\" " +
			"wsp:rsidRDefault=\"003D09F2\" wsp:rsidP=\"003D09F2\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【文件描述】　　　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"003D09F2\" " +
			"wsp:rsidRDefault=\"003D09F2\" wsp:rsidP=\"003D09F2\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/>" +
			"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【文件檔名】　　　　　　</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" " +
			"wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>【中文本原始檔】　　　　　　　</w:t></w:r><w:r><w:rPr>" +
			"<w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>#seq#-desc.doc</w:t></w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\">" +
			"<w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r>" +
			"<w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t></w:t>" +
			"</w:r></w:p>" +
			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【本申請書所檢送之PDF檔或影像檔與原本或正本相同】</w:t></w:r></w:p>" +

			"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" " +
			"w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" " +
			"w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>" +
			"【申請人已詳閱申請須知所定個人資料保護注意事項-並已確認本申請案之附件-除基本資料表-委任書外-不包含應予保密之個人資料-其載有個人資料者-同意智慧財產局提供任何人以自動化或非自動化之方式閱覽或抄錄或攝影或影印.】</w:t></w:r></w:p>";
		}

		//頁尾
		private string DocFooter() {
			return
			"<w:sectPr wsp:rsidR=\"000940B1\"><w:ftr w:type=\"odd\">" +
			"<w:p wsp:rsidR=\"000940B1\" wsp:rsidRDefault=\"000940B1\" wsp:rsidP=\"000940B1\">" +
			"<w:pPr><w:pStyle w:val=\"a7\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>第</w:t></w:r><w:fldSimple w:instr=\" PAGE   \\* MERGEFORMAT \"><w:r wsp:rsidR=\"00DE2999\"><w:rPr><w:noProof/></w:rPr>" +
			"<w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>頁，共</w:t>" +
			"</w:r><w:fldSimple w:instr=\" SECTIONPAGES  \\* MERGEFORMAT \"><w:r wsp:rsidR=\"00DE2999\"><w:rPr><w:noProof/></w:rPr><w:t>3</w:t></w:r>" +
			"</w:fldSimple><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>頁</w:t></w:r><w:r><w:rPr><w:rFonts " +
			"w:hint=\"fareast\"/></w:rPr><w:t>(</w:t></w:r><w:r wsp:rsidRPr=\"000940B1\"><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr><w:t>發明專利申請書</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/></w:rPr><w:t>)</w:t></w:r></w:p></w:ftr><w:pgSz " +
			"w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1134\" w:header=\"851\" w:footer=\"992\" " +
			"w:gutter=\"0\"/><w:cols w:space=\"425\"/><w:docGrid w:type=\"lines\" w:line-pitch=\"360\"/></w:sectPr>";
		}
		
		//頁尾+換頁
		private string DocNewPageFooter() {
			return
			"<w:p wsp:rsidR=\"000940B1\" " +
			"wsp:rsidRDefault=\"00397406\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing " +
			"w:line=\"360\" w:line-rule=\"at-least\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font wx:val=\"新細明體\"/>" +
			"</w:rPr>" + DocFooter() + "</w:pPr><w:r><w:rPr><w:rFonts " +
			"w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
			"</w:r></w:p>";
		}

		//結尾
		private string DocTail() {
			return "</w:body></w:wordDocument>";
		}
</script>
