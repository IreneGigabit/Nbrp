<%@ Application Language="C#" %>
<%@Import Namespace = "System.Data"%>
<%@Import Namespace = "System.Data.SqlClient"%>

<script runat="server">

    void Application_Start(object sender, EventArgs e) 
    {
        // 應用程式啟動時執行的程式碼
        //Application["CasePath"] = "\\\\web08\\Data$\\document\\XAccount";
        Application["CasePath"] = Server.MapPath("~/upload");
        Application["uploadDir"] = "nacc";
        Application["uploadPath"] = Server.MapPath("~/upload");
		Application["rootPath"] = "http://localhost/NACC";
        Application["photoPath"] = "D:\\Data\\Document\\NACC\\Temp\\images";
		Application["MailServer"] = "localhost";
		Application["MailAddr"] = "front.desk@my-farm.com.tw";
		Application["CRSmail"] = "customer.service@my-farm.com.tw";
		Application["farmTel"] = "049-2821418";
		Application["farmFax"] = "049-2821492";
		Application["farmAddr"] = "南投縣水里鄉上安村安田路9號";
		Application["MasilUID"] = "uuuu";
		Application["MailPWD"] = "pppp";
		Application["LoginImg"] = "NO";

		Application["P12File"] = Server.MapPath("~/inc") + "\\API-Project-f47f3742fc86.p12";
		Application["JsonFile"] = Server.MapPath("~/inc") + "\\snoopy-calendar.json";
		Application["svrEmail"] = "mycalendarsvr@api-project-114451378278.iam.gserviceaccount.com";
		Application["usrEmail"] = "snop222@gmail.com";
		Application["gcTestUID"] = "kitty";
		//Application["CalendarID"] = "primary";
		Application["CalendarID"] = "u0khi7hf8pd2skn7rosg42973k@group.calendar.google.com";
		Application["CldrStop"] = "GO";
		
		//Application["P12File"] = Server.MapPath("~/inc") + "\\MyFarmCalendar-e1c3efaf13e9.p12";
		//Application["svrEmail"] = "myfarm-52@myfarmcalendar.iam.gserviceaccount.com";
		//Application["usrEmail"] = "front.desk@my-farm.com.tw";
		//Application["gcTestUID"] = "my003";
		//Application["CalendarID"] = "primary";

		//SqlConnection cnn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring"].ToString());
		//string SQL = "SELECT * FROM tbCode WHERE vcType = 'Application'";
		//SqlCommand cmd = new SqlCommand(SQL, cnn);
		//cnn.Open();
		//SqlDataReader dr = cmd.ExecuteReader();
		//string sDefStr = "";
		////string sSpeCtrl = "";
		//while (dr.Read()) {
		//	sDefStr = (dr["vcRef1"] == DBNull.Value) ? dr["nvRef4"].ToString() : dr["vcRef1"].ToString();
		//	Application[dr["vcNo"].ToString()] = sDefStr;
		//}
		//cnn.Close();
	}
    
    void Application_End(object sender, EventArgs e) {
        //  應用程式關閉時執行的程式碼

    }
        
    void Application_Error(object sender, EventArgs e)  { 
        // 發生未處理錯誤時執行的程式碼
		// for IIS 6.0 (Windows 2003 Server)
		//Exception exObj = Server.GetLastError();
		//if (exObj is HttpUnhandledException)
		//{
		//	int hcode = ((HttpUnhandledException)exObj).GetHttpCode();
		//	if (hcode == 500) Server.Transfer("~/500-error.aspx");
		//}
	}

    void Session_Start(object sender, EventArgs e) {
        // 啟動新工作階段時執行的程式碼
        Session["ODBCDSN"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring"].ToString();
        Session["NACC"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring"].ToString();
        Session["ACCOUNT"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLAccount"].ToString();
        Session["MACCOUNT"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLMAccount"].ToString();
        Session["CUST"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLCust"].ToString();
		Session["SysCtrl"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring1"].ToString();       
        Session["BranchOLDB"] = System.Configuration.ConfigurationManager.ConnectionStrings["ODBCBranchCnnstringTest"].ToString();
		Session["HeadOLDB"] = System.Configuration.ConfigurationManager.ConnectionStrings["ODBCHeadCnnstringTest"].ToString();
		Session["imarraccount"] = System.Configuration.ConfigurationManager.ConnectionStrings["maccount"].ToString();//智產會計系統
		Session["btbrtdb"] = System.Configuration.ConfigurationManager.ConnectionStrings["dev_btbrtdb"].ToString();//案件資料使用
		Session["debit"] = "";//抓資料使用
        Session["Password"] = false;
        Session["UserID"] = Session.SessionID.ToString();
		Session["UserName"] = "";
		Session["UserGrp"] = "";
        Session["fSQL"] = "";
		Session["CaptchaImageText"] = "";
		Session["ExtMsg"] = "";
		Session["CustPasswd"] = false;
		Session["SCodeID"] = "";
		Session["CustID"] = "";
		Session["CustName"] = "";
		Session["CustEmail"] = "";        
		Session["Loc"] = "";
		Session["SvrName"] = "SIF02";
		Session["SvrName1"] = "SIF02";
		Session["Mobile"] = "Auto";
        Session["Syscode"] = "NAccount";
        Session["AccSvr"] = "web02";
        //正式
        Session["QRcode"] = "16618156"; //電子發票平台QRcode的密碼
        Session["QRGenKey"] = "E747993D0AF3D0199A7A9A56DABC106F"; //利用genKey.bat轉成Base64的QRcode碼
	}

    void Session_End(object sender, EventArgs e) 
    {
        // 工作階段結束時執行的程式碼。 
        // 注意: 只有在 Web.config 檔將 sessionstate 模式設定為 InProc 時，
        // 才會引發 Session_End 事件。如果將工作階段模式設定為 StateServer 
        // 或 SQLServer，就不會引發這個事件。

    }
       
</script>
