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
		
		in_scode = (Request["in_scode"] ?? "n100").ToString();//n100
		in_no = (Request["in_no"] ?? "20170103001").ToString();//20170103001
		branch = (Request["branch"] ?? "N").ToString();//N
		try {
			//Response.Write(("&#153706;瑄&#153706;").ToXmlUnicode());
			//Response.Write(Convert.ToChar((int)153706));
			string dbSession = (Session["btbrtdb"] ?? "Server=web08;Database=sindbs;User ID=web_usr;Password=web1823").ToString();
			//ipoRpt = new IpoReport(dbSession, in_scode, in_no, branch);
			ipoRpt = new IPOReport();
			WordOut();
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		Dictionary<string, string> _TemplateFileList = new Dictionary<string, string>();
		//_TemplateFileList.Add("apply", Server.MapPath("~/ReportTemplate/FE9團體標章註冊申請書.docx"));
		//_TemplateFileList.Add("base", Server.MapPath("~/ReportTemplate/00基本資料表.docx"));
		_TemplateFileList.Add("desc", Server.MapPath("~/ReportTemplate/說明書/01發明說明書IE_1.docx"));
		ipoRpt.CloneFromFile(_TemplateFileList, false);

		ipoRpt.ReplaceBookmark("cappl_name", "中文1");
		ipoRpt.ReplaceBookmark("eappl_name", "英文1");
		ipoRpt.ReplaceBookmark("cappl_name1", "英文2");
		ipoRpt.ReplaceBookmark("eappl_name1", "英文2");

		ipoRpt.Flush("[團體標章註冊申請書]-NT66824.docx");
	}
</script>
