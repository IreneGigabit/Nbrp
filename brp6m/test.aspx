<%@ Page Language="C#" %>

<!DOCTYPE html>

<script runat="server">
	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.Write("pageonload1.." + Session["btbrtdb"] + "<HR>");
		if (Request["branch"] != null) {
			Session["SeBranch"] = Request["branch"].ToString();
		}
		Global.StartSession();
		Response.Write("pageonload2.." + Session["btbrtdb"] + "<HR>");
		//Session.Abandon();
	}
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
</html>
