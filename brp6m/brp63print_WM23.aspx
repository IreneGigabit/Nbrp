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
                ReportCode = "WM23",
            }.Init();
            WordOut();
        }
        finally {
            if (ipoRpt != null) ipoRpt.Close();
        }
    }

    protected void WordOut() {
        Dictionary<string, string> _tplFile = new Dictionary<string, string>();
        _tplFile.Add("doc", Server.MapPath("~/ReportTemplate/申請書/修正申復理由書WM23.docx"));
        ipoRpt.CloneFromFile(_tplFile, false);

        string docFileName = ipoRpt.Seq + "-AllegationDocument.docx";
        if (Request["wordname"].ToString() == "RR11") {
            docFileName = ipoRpt.Seq + "-reason.docx";
        }
        
        ipoRpt.Flush(docFileName);
    }
</script>
