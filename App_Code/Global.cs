using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.SessionState;

public class Global : HttpApplication {
	void Application_Start(object sender, EventArgs e) {
		// 應用程式啟動時執行的程式碼
	}

	void Session_Start(object sender, EventArgs e) {
		// 啟動新工作階段時執行的程式碼
		StartSession();
	}

	public static void StartSession() {
		HttpContext.Current.Session["ODBCDSN"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring"].ToString();
		HttpContext.Current.Session["NACC"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring"].ToString();
		HttpContext.Current.Session["ACCOUNT"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLAccount"].ToString();
		HttpContext.Current.Session["MACCOUNT"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLMAccount"].ToString();
		HttpContext.Current.Session["CUST"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLCust"].ToString();
		HttpContext.Current.Session["SysCtrl"] = System.Configuration.ConfigurationManager.ConnectionStrings["SQLcnnstring1"].ToString();
		HttpContext.Current.Session["BranchOLDB"] = System.Configuration.ConfigurationManager.ConnectionStrings["ODBCBranchCnnstringTest"].ToString();
		HttpContext.Current.Session["HeadOLDB"] = System.Configuration.ConfigurationManager.ConnectionStrings["ODBCHeadCnnstringTest"].ToString();
		HttpContext.Current.Session["imarraccount"] = System.Configuration.ConfigurationManager.ConnectionStrings["maccount"].ToString();//智產會計系統
		//案件系統
		if (HttpContext.Current.Request.ServerVariables["HTTP_HOST"].ToString().ToUpper() == "WEB08") {
			//開發環境
			HttpContext.Current.Session["btbrtdb"] = System.Configuration.ConfigurationManager.ConnectionStrings["dev_btbrtdb"].ToString();
		} else if (HttpContext.Current.Request.ServerVariables["HTTP_HOST"].ToString().ToUpper() == "WEB10") {
			//使用者測試環境
			HttpContext.Current.Session["btbrtdb"] = System.Configuration.ConfigurationManager.ConnectionStrings["test_" + HttpContext.Current.Session["SeBranch"].ToString() + "_btbrtdb"].ToString();
		} else {
			//正式環境
			HttpContext.Current.Session["btbrtdb"] = System.Configuration.ConfigurationManager.ConnectionStrings["prod_" + HttpContext.Current.Session["SeBranch"].ToString() + "_btbrtdb"].ToString();
		}
		HttpContext.Current.Response.Write("sessionstart.." + DateTime.Now.ToString() + " - " + HttpContext.Current.Session["SeBranch"] + " - " + HttpContext.Current.Session["btbrtdb"] + "<HR>");
		HttpContext.Current.Session["debit"] = "";//抓資料使用
		HttpContext.Current.Session["Password"] = false;
		HttpContext.Current.Session["UserID"] = HttpContext.Current.Session.SessionID.ToString();
		HttpContext.Current.Session["UserName"] = "";
		HttpContext.Current.Session["UserGrp"] = "";
		HttpContext.Current.Session["fSQL"] = "";
		HttpContext.Current.Session["CaptchaImageText"] = "";
		HttpContext.Current.Session["ExtMsg"] = "";
		HttpContext.Current.Session["CustPasswd"] = false;
		HttpContext.Current.Session["SCodeID"] = "";
		HttpContext.Current.Session["CustID"] = "";
		HttpContext.Current.Session["CustName"] = "";
		HttpContext.Current.Session["CustEmail"] = "";
		HttpContext.Current.Session["Loc"] = "";
		HttpContext.Current.Session["SvrName"] = "SIF02";
		HttpContext.Current.Session["SvrName1"] = "SIF02";
		HttpContext.Current.Session["Mobile"] = "Auto";
		HttpContext.Current.Session["Syscode"] = "NAccount";
		HttpContext.Current.Session["AccSvr"] = "web02";
		//正式
		HttpContext.Current.Session["QRcode"] = "16618156"; //電子發票平台QRcode的密碼
		HttpContext.Current.Session["QRGenKey"] = "E747993D0AF3D0199A7A9A56DABC106F"; //利用genKey.bat轉成Base64的QRcode碼
	}
}
