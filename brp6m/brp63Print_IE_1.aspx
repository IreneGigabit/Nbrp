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
		try {
			ipoRpt = new IPOReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch);
			WordOut();
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		Dictionary<string, string> _TemplateFileList = new Dictionary<string, string>();
		_TemplateFileList.Add("desc", Server.MapPath("~/ReportTemplate/說明書/01發明說明書IE_1.docx"));
		ipoRpt.CloneFromFile(_TemplateFileList, false);

		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			ipoRpt.ReplaceText("#cappl_name#", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
			ipoRpt.ReplaceText("#eappl_name#", dmp.Rows[0]["eappl_name"].ToString().ToXmlUnicode());
		}

		ipoRpt.Flush(ipoRpt.Seq + "-desc.docx");
	}
</script>
