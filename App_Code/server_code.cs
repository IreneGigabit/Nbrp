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
	#region Left
	public static string Left(this string str, int ln) {
		string sret = str.Substring(0, Math.Min(ln, str.Length));
		return sret;
	}
	#endregion

	#region Right
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
	#endregion

	#region ToXmlUnicode - 將&#nnnn;轉成word用格式
	/// <summary>
	/// 將&amp;#nnnn;轉成word用格式
	/// </summary>
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
	#endregion

	#region ToBig5 - 將難字轉成&#nnnn;
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
	#endregion
}