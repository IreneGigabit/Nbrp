﻿<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>

<script runat="server">
	protected string in_scode = "";
	protected string in_no = "";
	protected string branch = "";

	protected IPOReport ipoRpt = null;

	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;
		Response.Clear();

		in_scode = (Request["in_scode"] ?? "").ToString();//n100
		in_no = (Request["in_no"] ?? "").ToString();//20170103001
		branch = (Request["branch"] ?? "").ToString();//N
		string rectitle = (Request["receipt_title"] ?? "").ToString();//N

		try {
			//電子收據第2階段上線後要廢除RectitleTitle參數
			ipoRpt = new IPOReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch, rectitle);
			WordOut();
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		string applyFile="";
		if (Request["wordname"].ToString() == "UH1_1" || Request["wordname"].ToString() == "UH1_2") {
			applyFile= "27發明專利改請新型專利申請書UH1.docx";
		} else if (Request["wordname"].ToString() == "UH1_3" || Request["wordname"].ToString() == "UH1_4") {
			applyFile= "20設計專利改請新型專利申請書UH1.docx";
		}
		Dictionary<string, string> _tplFile = new Dictionary<string, string>();
		_tplFile.Add("apply", Server.MapPath("~/ReportTemplate") + @"\"+applyFile);
		_tplFile.Add("base", Server.MapPath("~/ReportTemplate") + @"\00基本資料表.docx");
		ipoRpt.CloneFromFile(_tplFile, true);

		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			//標題區塊
			ipoRpt.CopyBlock("b_title");
			//原申請案號
			if (dmp.Rows[0]["change_no"].ToString() != "") {
				ipoRpt.ReplaceBookmark("apply_no", dmp.Rows[0]["change_no"].ToString());
			} else {
				ipoRpt.ReplaceBookmark("apply_no", dmp.Rows[0]["apply_no"].ToString());
			}
			//事務所或申請人案件編號
			ipoRpt.ReplaceBookmark("seq", ipoRpt.Seq + "-" + dmp.Rows[0]["scode1"].ToString());
			//中文名稱 / 英文名稱
			ipoRpt.ReplaceBookmark("cappl_name", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
			ipoRpt.ReplaceBookmark("eappl_name", dmp.Rows[0]["eappl_name"].ToString().ToXmlUnicode());
			//申請人
			using (DataTable dtAp = ipoRpt.Apcust) {
				for (int i = 0; i < dtAp.Rows.Count; i++) {
					ipoRpt.CopyBlock("b_apply");
					ipoRpt.ReplaceBookmark("apply_num", (i + 1).ToString());
					ipoRpt.ReplaceBookmark("ap_country", dtAp.Rows[i]["Country_name"].ToString());
					ipoRpt.ReplaceBookmark("ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
					ipoRpt.ReplaceBookmark("ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
					ipoRpt.ReplaceBookmark("ap_cname", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
					ipoRpt.ReplaceBookmark("ap_ename", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				}
			}
			//代理人
			ipoRpt.CopyBlock("b_agent");
			using (DataTable dtAgt = ipoRpt.Agent) {
				ipoRpt.ReplaceBookmark("agt_name1", dtAgt.Rows[0]["agt_name1"].ToString().Trim());
				ipoRpt.ReplaceBookmark("agt_name2", dtAgt.Rows[0]["agt_name2"].ToString().Trim());
			}
			//新型創作人
			using (DataTable dtAnt = ipoRpt.Ant) {
				for (int i = 0; i < dtAnt.Rows.Count; i++) {
					ipoRpt.CopyBlock("b_ant");
					ipoRpt.ReplaceBookmark("ant_num", (i + 1).ToString());
					ipoRpt.ReplaceBookmark("ant_country", dtAnt.Rows[i]["Country_name"].ToString());
					ipoRpt.ReplaceBookmark("ant_cname", dtAnt.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
					ipoRpt.ReplaceBookmark("ant_ename", dtAnt.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				}
			}
			//主張優惠期
			string exh_date = "";
			if (dmp.Rows[0]["exhibitor"].ToString() == "Y") {//參展或發表日期填入表中的發生日期
				if (dmp.Rows[0]["exh_date"] != System.DBNull.Value && dmp.Rows[0]["exh_date"] != null) {
					exh_date = Convert.ToDateTime(dmp.Rows[0]["exh_date"]).ToString("yyyy/MM/dd");
				}
			}
			ipoRpt.CopyBlock("b_exh");
			ipoRpt.ReplaceBookmark("exh_date", exh_date);
			
			//文本資訊/繳費資訊
			ipoRpt.CopyBlock("b_content");
			ipoRpt.ReplaceBookmark("receipt_name", ipoRpt.RectitleName);
			//附送書件
			ipoRpt.CopyReplaceBlock("b_attach", "#seq#", ipoRpt.Seq);
			//具結
			ipoRpt.CopyBlock("b_sign");

			bool baseflag = true;//是否產生基本資料表
			ipoRpt.CopyPageFoot("apply", baseflag);//申請書頁尾
			if (baseflag) {
				ipoRpt.AppendBaseData("base", "新型創作人");//產生基本資料表
			}
		}

		if (Request["wordname"].ToString() == "UH1_4") {
			ipoRpt.Flush(Request["se_scode"] + "-UH1_4_change_form.docx");//UH1_4設計專利改請新型專利申請書(含說明書)
		} else if (Request["wordname"].ToString() == "UH1_3") {
			ipoRpt.Flush(Request["se_scode"] + "-UH1_3_change_form.docx");//UH1_3設計專利改請新型專利申請書
		} else if (Request["wordname"].ToString() == "UH1_2") {
			ipoRpt.Flush(Request["se_scode"] + "-UH1_2_change_form.docx");//UH1_2發明專利改請新型專利申請書(含說明書)
		} else if (Request["wordname"].ToString() == "UH1_1") {
			ipoRpt.Flush(Request["se_scode"] + "-UH1_1_change_form.docx");//UH1_1發明專利改請新型專利申請書
		} else {
			ipoRpt.Flush(Request["se_scode"] + "-UH1_patent_form.docx");
		}
		ipoRpt.SetPrint();
	}
</script>
