using System;
using System.Collections.Generic;
using System.Web;
using System.IO;
using System.Threading;
using System.Text;

/// <summary>
/// server_code 的摘要描述
/// </summary>
public static class server_code {
	static string SiteDir = HttpContext.Current.Request.PhysicalApplicationPath;

	private static object lockObject = new object();

	#region 產出Exception log
	public static void exceptionLog(Exception ex) {
		exceptionLog(ex, "");
	}

	public static void exceptionLog(Exception ex, string sql) {
		string Message = "";

		if (sql != "") {
			Message = "發生錯誤的網頁:{0}\n錯誤訊息:{1}\nSQL:\n{2}\n堆疊內容:\n{3}\n";
			Message = String.Format(Message, HttpContext.Current.Request.Path, ex.GetBaseException().Message, sql, ex.StackTrace);
		} else {
			Message = "發生錯誤的網頁:{0}\n錯誤訊息:{1}\n堆疊內容:\n{3}\n";
			Message = String.Format(Message, HttpContext.Current.Request.Path, ex.GetBaseException().Message, ex.StackTrace);
		}

		writeLog(Message);
	}
	#endregion

	#region 寫入log
	public static void writeLog(string detailDesc) {
		Monitor.Enter(lockObject);
		StreamWriter sw = null;
		try {
			string Path = string.Format(@"{0}Logs\{1}\", SiteDir, DateTime.Now.ToString("yyyy"));
			if (!Directory.Exists(Path)) {
				Directory.CreateDirectory(Path);
			}

			string logFile = Path + DateTime.Now.ToString("yyyyMMdd") + ".txt";
			string fullText = "";
			if (detailDesc == "")
				fullText = "";
			else
				fullText = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\n" + detailDesc;

			sw = File.AppendText(logFile);
			sw.WriteLine(fullText);
			sw.Flush();
		}
		finally {
			Monitor.Exit(lockObject);
			if (sw != null) sw.Close();
		}
	}
	#endregion

	public static string Left(this string str, int ln) {
		string sret = str.Substring(0, Math.Min(ln, str.Length));
		return sret;
	}

	public static string Right(this string str, int ln) {
		//string sret = s.Substring(str.Length - ln, ln);
		//return sret;
		ln = Math.Max(ln, 0);
		if (str.Length > ln) {
			return str.Substring(str.Length - ln, ln);
		} else {
			return str;
		}
	}

	public static string ToXmlUnicode(this string str) {
		return str.ToXmlUnicode(false);
	}
	public static string ToXmlUnicode(this string str, bool isEng) {
		str = HttpUtility.HtmlDecode(str);
		foreach (System.Text.RegularExpressions.Match m
			in System.Text.RegularExpressions.Regex.Matches(str, "&#(?<ncr>\\d+?);"))
			str = str.Replace(m.Value, Convert.ToChar(int.Parse(m.Groups["ncr"].Value)).ToString());
		//str = str.Replace("&", "&amp;");
		//str = str.Replace("<", "&lt;");
		if (isEng) {//防止英文欄位只能半型
			str = str.Replace("’", "'");
			str = str.Replace("＆", "&");
		}
		//ret=str.Replace(">","&gt;");
		//ret=str.Replace("'","&apos;");
		//ret=str.Replace("""","&quot;");

		return str.Trim();
	}
	/// <summary>
	/// 將難字轉成&amp;#nnnn;
	/// </summary>
	public static string ToBig5(this string str) {
		StringBuilder sb = new StringBuilder();
		Encoding big5 = Encoding.GetEncoding("big5");
		foreach (char c in str) {
			string cInBig5 = big5.GetString(big5.GetBytes(new char[] { c }));
			if (c != '?' && cInBig5 == "?")
				sb.AppendFormat("&#{0};", Convert.ToInt32(c));
			else
				sb.Append(c);
		}
		return sb.ToString();
	}
}