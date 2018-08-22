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
		try {
            ipoRpt = new IPOReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch)
            {
                ReportCode = "UE_1",
            }.Init();
			WordOut();
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		Dictionary<string, string> _tplFile = new Dictionary<string, string>();
		_tplFile.Add("desc", Server.MapPath("~/ReportTemplate/說明書/02新型說明書UE_1.docx"));
		ipoRpt.CloneFromFile(_tplFile, false);
		
		string docFileName = ipoRpt.Seq + "-desc.docx";

		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			ipoRpt.ReplaceText("#cappl_name#", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
			ipoRpt.ReplaceText("#eappl_name#", dmp.Rows[0]["eappl_name"].ToString().ToXmlUnicode());
		}

		ipoRpt.Flush(docFileName);
	}
</script>
