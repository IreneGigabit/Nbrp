<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>

<script runat="server">
	protected IPOReport ipoRpt = null;

	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;
		Response.Clear();

        string in_scode = (Request["in_scode"] ?? "").ToString();//n100
        string in_no = (Request["in_no"] ?? "").ToString();//20170103001
        string branch = (Request["branch"] ?? "").ToString();//N
		string rectitle = (Request["receipt_title"] ?? "").ToString();//N

		try {
            ipoRpt = new IPOReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch, rectitle)
            {
                ReportCode = "DE1",
            }.Init();
			WordOut();
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		Dictionary<string, string> _tplFile = new Dictionary<string, string>();
		_tplFile.Add("apply", Server.MapPath("~/ReportTemplate/申請書/03設計專利申請書DE1.docx"));
		_tplFile.Add("base", Server.MapPath("~/ReportTemplate/申請書/00基本資料表.docx"));
		ipoRpt.CloneFromFile(_tplFile, true);

		string docFileName = ipoRpt.Seq + "-設計.docx";
			
		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			//標題區塊
			ipoRpt.CopyBlock("b_title");
			//事務所或申請人案件編號
			ipoRpt.ReplaceBookmark("seq", ipoRpt.Seq + "-" + dmp.Rows[0]["scode1"].ToString());
			//中文設計名稱 / 英文設計名稱
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
			//設計人
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

			//主張優先權
			using (DataTable dtPrior = ipoRpt.Prior) {
				for (int i = 0; i < dtPrior.Rows.Count; i++) {
					string prior_date = "";
					if (dtPrior.Rows[i]["prior_date"] != System.DBNull.Value && dtPrior.Rows[i]["prior_date"] != null) {
						prior_date = Convert.ToDateTime(dtPrior.Rows[i]["prior_date"]).ToString("yyyy/MM/dd");
					}
					ipoRpt.CopyBlock("b_prior1");
					ipoRpt.ReplaceBookmark("prior_num", (i + 1).ToString());
					ipoRpt.ReplaceBookmark("prior_country", dtPrior.Rows[i]["Country_name"].ToString());
					ipoRpt.ReplaceBookmark("prior_date", prior_date);
					ipoRpt.ReplaceBookmark("prior_no", dtPrior.Rows[i]["prior_no"].ToString());
				}
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
				ipoRpt.AppendBaseData("base", "設計人");//產生基本資料表
			}
		}

		ipoRpt.Flush(docFileName);
		ipoRpt.SetPrint();
	}
</script>
