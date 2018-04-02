<%@ Page Language="C#" %>
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
		Dictionary<string, string> _tplFile = new Dictionary<string, string>();
		_tplFile.Add("apply", Server.MapPath("~/ReportTemplate/申請書/35[專簡A]專利申請延展指定期間申請書WMB.docx"));
		_tplFile.Add("base", Server.MapPath("~/ReportTemplate/申請書/00基本資料表.docx"));
		ipoRpt.CloneFromFile(_tplFile, true);
		
		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			//標題區塊
			ipoRpt.CopyBlock("b_title");
			//專利類別
			string Case1nm = "";
			switch (dmp.Rows[0]["CASE1"].ToString().Substring(0, 2)) {
				case "IG":
					Case1nm = "發明";
					break;
				case "UG":
					Case1nm = "新型";
					break;
				case "DG":
					Case1nm = "設計";
					break;
			}
			ipoRpt.ReplaceBookmark("case1nm", Case1nm);
			//原申請案號
			if (dmp.Rows[0]["change_no"].ToString() != "") {
				ipoRpt.ReplaceBookmark("apply_no", dmp.Rows[0]["change_no"].ToString());
			} else {
				ipoRpt.ReplaceBookmark("apply_no", dmp.Rows[0]["apply_no"].ToString());
			}
			//事務所或申請人案件編號
			ipoRpt.ReplaceBookmark("seq", ipoRpt.Seq + "-" + dmp.Rows[0]["scode1"].ToString());
			//專利名稱
			ipoRpt.ReplaceBookmark("cappl_name", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
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
				ipoRpt.AppendBaseData("base", "");//產生基本資料表
			}
		}

		ipoRpt.Flush(Request["se_scode"] + "-WMB_form.docx");
		ipoRpt.SetPrint();
	}
</script>
