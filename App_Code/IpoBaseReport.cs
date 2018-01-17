using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Text;

/// <summary>
/// 產生智慧局電子申請書-基本資料表用
/// </summary>
public class IPOBaseReport : IpoReportXml {
	public StringBuilder sbXml = new StringBuilder();

	public IPOBaseReport(string connStr,string in_scode,  string in_no) {
		this._connStr = connStr;
		this._in_no = in_no;
		this._in_scode = in_scode;
		this._conn = new DBHelper(connStr, false).Debug(false);
	}

	public IPOBaseReport Close() {
		_conn.Dispose();
		return this;
	}

	public IPOBaseReport Build(string type_string) {
		string xml="";
		//抬頭
		sbXml.Append(DocBody_11());

		//產生 基本資料表-申請人
		using (DataTable dtAp = GetBaseApcust()) {
			for (int i = 0; i < dtAp.Rows.Count; i++) {
				xml = Apcust_data_1();
				xml = xml.Replace("#apply_num#", (i + 1).ToString());
				xml = xml.Replace("#ap_country#", dtAp.Rows[i]["Country_name"].ToString());
				xml = xml.Replace("#ap_class#", dtAp.Rows[i]["apclass_name"].ToString());
				sbXml.Append(xml);
				if (dtAp.Rows[i]["ap_country"].ToString() == "T") {
					xml = Apcust_data_1_2();
					xml = xml.Replace("#apcust_no#", dtAp.Rows[i]["apcust_no"].ToString().Trim());
					sbXml.Append(xml);
				}
				xml = Apcust_data_1_3();
				xml = xml.Replace("#ap_cname1_title#", dtAp.Rows[i]["Title_cname"].ToString());
				xml = xml.Replace("#ap_country#", dtAp.Rows[i]["Country_name"].ToString());
				xml = xml.Replace("#ap_ename1_title#", dtAp.Rows[i]["Title_ename"].ToString());
				xml = xml.Replace("#ap_cname1#", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ap_ename1#", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode());

				xml = xml.Replace("#ap_zip#", dtAp.Rows[i]["ap_zip"].ToString().Trim());
				xml = xml.Replace("#ap_addr1#", dtAp.Rows[i]["ap_addr1"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ap_addr2#", dtAp.Rows[i]["ap_addr2"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ap_eaddr1#", dtAp.Rows[i]["ap_eaddr1"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ap_eaddr2#", dtAp.Rows[i]["ap_eaddr2"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ap_eaddr3#", dtAp.Rows[i]["ap_eaddr3"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ap_eaddr4#", dtAp.Rows[i]["ap_eaddr4"].ToString().ToXmlUnicode());

				xml = xml.Replace("#ap_crep#", dtAp.Rows[i]["ap_crep"].ToString().Trim());
				xml = xml.Replace("#ap_erep#", dtAp.Rows[i]["ap_erep"].ToString().ToXmlUnicode());
				sbXml.Append(xml);
			}
		}

		//產生 基本資料表-代理人
		using (DataTable dtAgt = GetBaseAgt()) {
			xml = Agt_data();//代理人1
			xml = xml.Replace("#agt_num#", "1");
			xml = xml.Replace("#agt_idno#", dtAgt.Rows[0]["agt_idno1"].ToString().Trim());
			xml = xml.Replace("#agt_id#", dtAgt.Rows[0]["agt_id1"].ToString().Trim());
			xml = xml.Replace("#agt_name#", dtAgt.Rows[0]["agt_name1"].ToString().Trim());
			xml = xml.Replace("#agt_zip#", dtAgt.Rows[0]["agt_zip"].ToString().Trim());
			xml = xml.Replace("#agt_addr#", dtAgt.Rows[0]["agt_addr"].ToString().Trim());
			xml = xml.Replace("#agt_tel#", dtAgt.Rows[0]["agt_tel"].ToString().Trim());
			xml = xml.Replace("#agt_fax#", dtAgt.Rows[0]["agt_fax"].ToString().Trim());
			sbXml.Append(xml);
			xml = Agt_data();//代理人2
			xml = xml.Replace("#agt_num#", "2");
			xml = xml.Replace("#agt_idno#", dtAgt.Rows[0]["agt_idno2"].ToString().Trim());
			xml = xml.Replace("#agt_id#", dtAgt.Rows[0]["agt_id2"].ToString().Trim());
			xml = xml.Replace("#agt_name#", dtAgt.Rows[0]["agt_name2"].ToString().Trim());
			xml = xml.Replace("#agt_zip#", dtAgt.Rows[0]["agt_zip"].ToString().Trim());
			xml = xml.Replace("#agt_addr#", dtAgt.Rows[0]["agt_addr"].ToString().Trim());
			xml = xml.Replace("#agt_tel#", dtAgt.Rows[0]["agt_tel"].ToString().Trim());
			xml = xml.Replace("#agt_fax#", dtAgt.Rows[0]["agt_fax"].ToString().Trim());
			sbXml.Append(xml);
		}

		//產生 基本資料表-發明人
		using (DataTable dtAnt = GetBaseAnt()) {
			for (int i = 0; i < dtAnt.Rows.Count; i++) {
				xml = Ant_data_1();
				xml = xml.Replace("#ant_num#", type_string + (i + 1).ToString());
				xml = xml.Replace("#ant_country#", dtAnt.Rows[i]["Country_name"].ToString());
				sbXml.Append(xml);

				if (dtAnt.Rows[i]["ant_country"].ToString() == "T") {
					xml = Ant_data_1_1();
					xml = xml.Replace("#ant_id#", dtAnt.Rows[i]["ant_id"].ToString().Trim());
					sbXml.Append(xml);
				}

				xml = Ant_data_1_2();
				xml = xml.Replace("#ant_cname#", dtAnt.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				xml = xml.Replace("#ant_ename#", dtAnt.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				sbXml.Append(xml);
			}
		}

		sbXml.Append(DocFooter());
		//sw.Append(DocTail());
		return this;
	}

	//空白行
	private string SpaceString() {
		return
		"<w:p wsp:rsidR=\"008515F6\" " +
	"wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/><wx:font " +
	"wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
	"<wx:font wx:val=\"新細明體\"/></w:rPr></w:r></w:p>";
	}

	//基本資料表 個人資料 
	private string DocBody_11() {
		return
		"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【基本資料】　　　</w:t>" +
		"</w:r></w:p><w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"008515F6\" wsp:rsidP=\"008515F6\"><w:pPr><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>【個人資料】　　　</w:t>" +
		"</w:r></w:p>";
	}

	//申請人
	private string Apcust_data_1() {
		return
		"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"00565151\" " +
		"wsp:rsidP=\"00565151\"><w:pPr><w:pStyle w:val=\"ac\"/><w:ind w:left=\"0\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【申請人#apply_num#】</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\">" +
		"<w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【國籍】　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"</w:rPr><w:t>#ap_country#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" " +
		"wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/>" +
		"</w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【身分種類】　　　　</w:t>" +
		"</w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/></w:rPr><w:t>#ap_class#</w:t></w:r></w:p>";
	}

	private string Apcust_data_1_2() {
		return
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/>" +
		"<w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr>" +
		"<w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/></w:rPr><w:t>#apcust_no#</w:t></w:r></w:p>";
	}

	private string Apcust_data_1_3() {
		return
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【#ap_cname1_title#】　　　　</w:t>" +
		"</w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/></w:rPr><w:t>#ap_cname1#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRPr=\"002F4BE0\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【#ap_ename1_title#】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/></w:rPr><w:t>#ap_ename1#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid " +
		"w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/>" +
		"</w:rPr><w:t>　　　【居住國】　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Times New " +
		"Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/><w:color w:val=\"000000\"/>" +
		"</w:rPr><w:t>#ap_country#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\">" +
		"<w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【郵遞區號】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Times New " +
		"Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/><w:color w:val=\"000000\"/>" +
		"</w:rPr><w:t>#ap_zip#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" " +
		"w:line-rule=\"auto\"/><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【中文地址】　　　　#ap_addr1##ap_addr2#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" " +
		"wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/>" +
		"<w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times " +
		"New Roman\"/><w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr>" +
		"<w:t>　　　【英文地址】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"Times New Roman\" " +
		"w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/><w:color w:val=\"000000\"/></w:rPr><w:t>#ap_eaddr1##ap_eaddr2##ap_eaddr3##ap_eaddr4#</w:t>" +
		"</w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
		"w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【代表人中文姓名】　#ap_crep#</w:t></w:r>" +
		"</w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
		"w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【代表人英文姓名】　#ap_erep#</w:t></w:r></w:p>" +
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
		"w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【法定代理人ID】　　</w:t></w:r>" +
		"</w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
		"w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【法定代理人中文姓名】</w:t></w:r></w:p>" + SpaceString();
	}

	//代理人
	private string Agt_data() {
		return
		"<w:p wsp:rsidR=\"00754829\" wsp:rsidRDefault=\"00754829\" wsp:rsidP=\"00754829\"><w:pPr><w:overflowPunct w:val=\"off\"/>" +
		"<w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/></w:pPr><w:r><w:rPr>" +
		"<w:rFonts w:hint=\"fareast\"/></w:rPr><w:t>　　【代理人#agt_num#】</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【證書字號】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts " +
		"w:ascii=\"Times New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New " +
		"Roman\"/><w:color w:val=\"000000\"/></w:rPr><w:t>#agt_idno#</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr>" +
		"<w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times " +
		"New Roman\"/><w:color w:val=\"000000\"/></w:rPr><w:t>#agt_id#</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"<w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【中文姓名】　　　　#agt_name#</w:t>" +
		"</w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
		"w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【郵遞區號】　　　　</w:t>" +
		"</w:r><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font " +
		"wx:val=\"Times New Roman\"/><w:color w:val=\"000000\"/></w:rPr><w:t>#agt_zip#</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\">" +
		"<w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【中文地址】　　　　#agt_addr#</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"Times " +
		"New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/><w:color " +
		"w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【電話】　　　　　　#agt_tel#</w:t></w:r>" +
		"</w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing " +
		"w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【傳真】　　　　　　#agt_fax#</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRPr=\"00FE23DF\" " +
		"wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:overflowPunct w:val=\"off\"/><w:autoSpaceDE w:val=\"off\"/><w:autoSpaceDN " +
		"w:val=\"off\"/><w:spacing w:line=\"360\" w:line-rule=\"at-least\"/><w:rPr><w:rFonts w:ascii=\"Calibri\" w:h-ansi=\"Calibri\" w:hint=\"fareast\"/>" +
		"<wx:font wx:val=\"Calibri\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【E-mail】　　　　　　</w:t>" +
		"</w:r><w:r wsp:rsidRPr=\"00EF5BA9\"><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r wsp:rsidRPr=\"00EF5BA9\"><w:instrText> HYPERLINK " +
		"\"mailto:siiplo@mail.saint-island.com.tw\" </w:instrText></w:r><w:r wsp:rsidRPr=\"00EF5BA9\"><w:fldChar w:fldCharType=\"separate\"/>" +
		"</w:r><w:r wsp:rsidRPr=\"00EF5BA9\"><w:t>siiplo@mail.saint-island.com.tw</w:t></w:r><w:r wsp:rsidRPr=\"00EF5BA9\"><w:fldChar w:fldCharType=\"end\"/>" +
		"</w:r></w:p>" + SpaceString();
	}

	//發明人
	private string Ant_data_1() {
		return
		"<w:p wsp:rsidR=\"008515F6\" wsp:rsidRDefault=\"00565151\" " +
		"wsp:rsidP=\"00565151\"><w:pPr><w:pStyle w:val=\"ac\"/><w:ind w:left=\"0\"/><w:rPr><w:rFonts w:ascii=\"新細明體\" w:h-ansi=\"新細明體\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"新細明體\" w:hint=\"fareast\"/>" +
		"<wx:font wx:val=\"新細明體\"/></w:rPr><w:t>　　【#ant_num#】</w:t></w:r></w:p><w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\">" +
		"<w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/>" +
		"<w:color w:val=\"000000\"/></w:rPr><w:t>　　　【國籍】　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Times New " +
		"Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/><w:color w:val=\"000000\"/>" +
		"</w:rPr><w:t>#ant_country#</w:t></w:r></w:p>";
	}

	//發明人
	private string Ant_data_1_1() {
		return
		"<w:p " +
		"wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" " +
		"w:line-rule=\"auto\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:h-ansi=\"Times New Roman\" w:cs=\"Times " +
		"New Roman\" w:hint=\"fareast\"/><wx:font wx:val=\"Times New Roman\"/><w:color w:val=\"000000\"/></w:rPr><w:t>#ant_id#</w:t></w:r></w:p>";
	}

	private string Ant_data_1_2() {
		return
		"<w:p wsp:rsidR=\"00123906\" wsp:rsidRDefault=\"00123906\" " +
		"wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:color w:val=\"000000\"/>" +
		"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【中文姓名】　　　　#ant_cname#</w:t></w:r>" +
		"</w:p><w:p wsp:rsidR=\"00754829\" wsp:rsidRPr=\"00123906\" wsp:rsidRDefault=\"00123906\" wsp:rsidP=\"00123906\"><w:pPr><w:snapToGrid " +
		"w:val=\"off\"/><w:spacing w:line=\"288\" w:line-rule=\"auto\"/><w:rPr><w:color w:val=\"000000\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/><w:color w:val=\"000000\"/></w:rPr><w:t>　　　【英文姓名】　　　　#ant_ename#</w:t></w:r></w:p>";
	}

	//頁尾
	private string DocFooter() {
		return
		"<w:sectPr wsp:rsidR=\"00397406\" wsp:rsidRPr=\"008515F6\" wsp:rsidSect=\"008515F6\">" +
		"<w:ftr w:type=\"odd\"><w:p wsp:rsidR=\"00A97328\" wsp:rsidRPr=\"00D01292\" wsp:rsidRDefault=\"00945317\" wsp:rsidP=\"00A97328\"><w:pPr>" +
		"<w:pStyle w:val=\"a7\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/>" +
		"</w:rPr><w:t>第</w:t></w:r><w:fldSimple w:instr=\" PAGE   \\* MERGEFORMAT \"><w:r wsp:rsidR=\"00A47B5D\"><w:rPr><w:noProof/></w:rPr>" +
		"<w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>頁，共</w:t>" +
		"</w:r><w:fldSimple w:instr=\" SECTIONPAGES  \\* MERGEFORMAT \"><w:r wsp:rsidR=\"00A47B5D\"><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>" +
		"</w:fldSimple><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr><w:t>頁</w:t></w:r><w:r><w:rPr><w:rFonts " +
		"w:hint=\"fareast\"/></w:rPr><w:t>(</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/><wx:font wx:val=\"新細明體\"/></w:rPr>" +
		"<w:t>基本資料表</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"fareast\"/></w:rPr><w:t>)</w:t></w:r></w:p></w:ftr><w:pgSz w:w=\"11906\" " +
		"w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1134\" w:header=\"851\" w:footer=\"992\" w:gutter=\"0\"/>" +
		"<w:pgNumType w:start=\"1\"/><w:cols w:space=\"425\"/><w:docGrid w:line-pitch=\"360\"/></w:sectPr>";
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
}