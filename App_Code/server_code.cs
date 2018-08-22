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
	#region 傳回字串，包含指定的從字串左邊的字元數 +static string Left(this string str, int ln)
	public static string Left(this string str, int ln) {
		string sret = str.Substring(0, Math.Min(ln, str.Length));
		return sret;
	}
	#endregion

	#region 傳回字串，包含指定的從字串右邊的字元數 +static string Right(this string str, int ln)
	public static string Right(this string str, int ln) {
		ln = Math.Max(ln, 0);
		if (str.Length > ln) {
			return str.Substring(str.Length - ln, ln);
		} else {
			return str;
		}
	}
	#endregion

	#region 將&#nnnn;轉成word用格式 +static string ToXmlUnicode(this string str)
	/// <summary>
	/// 將&amp;#nnnn;轉成word用格式
	/// </summary>
	public static string ToXmlUnicode(this string str) {
		return str.ToXmlUnicode(false);
	}
	#endregion

	#region 將&#nnnn;轉成word用格式 +static string ToXmlUnicode(this string str, bool isEng)
	/// <summary>
	/// 將&amp;#nnnn;轉成word用格式
	/// </summary>
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

	#region 將難字轉成&#nnnn; +static string ToBig5
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

	#region 截取字串,指定長度 +static string CutStr(this string str, int len)
	/// <summary>
	/// 截取字串,指定長度
	/// </summary>
	/// <param name="len">截取長度</param>
	/// <returns></returns>
	public static string CutStr(this string str, int len) {
		if (str == null || str.Length == 0 || len <= 0) {
			return string.Empty;
		}

		int orgLen = str.Length;

		int clen = 0;
		//計算要substr的長度
		while (clen < len && clen < orgLen) {
			//每遇到一個中文，則將目標長度減一。
			if ((int)str[clen] > 128) { len--; }
			clen++;
		}

		if (clen < orgLen) {
			return str.Substring(0, clen);
		} else {
			return str;
		}
	}
	//public static string CutStr(this string str, int totalWidth) {
	//	Encoding l_Encoding = Encoding.GetEncoding("big5", new EncoderExceptionFallback(), new DecoderReplacementFallback(""));
	//	byte[] strut8 = Encoding.Unicode.GetBytes(str);
	//	//byte[] strbig5 = Encoding.Convert(Encoding.Unicode, Encoding.GetEncoding("big5"), strut8);
	//	//byte[] strbig5 = Encoding.GetEncoding("big5").GetBytes(str);
	//	byte[] strbig5 = l_Encoding.GetBytes(str);
	//	return l_Encoding.GetString(strbig5, 0, totalWidth);
	//
	//	//Encoding l_Encoding = Encoding.GetEncoding("big5", new EncoderExceptionFallback(), new DecoderReplacementFallback(""));
	//	//byte[] l_byte = l_Encoding.GetBytes(str);
	//	//HttpContext.Current.Response.Write("CutStr=" + str + "(" + l_byte.Length + ")" + Environment.NewLine);
	//	//return l_Encoding.GetString(l_byte, 0, totalWidth);
	//}
	#endregion

	#region 字串靠右對齊 +static string PadLeftCHT(this string str, int totalWidth, char paddingChar)
	/// <summary>
	/// 字串靠右對齊，以指定的字元在左側補足長度，超過則截字(中文算2碼)。
	/// </summary>
	/// <param name="totalWidth">長度</param>
	/// <param name="paddingChar">替代字元</param>
	/// <returns></returns>
	public static string PadLeftCHT(this string str, int totalWidth, char paddingChar) {
		string sResult = str.CutStr(totalWidth);
		int orgLen = Encoding.GetEncoding("big5").GetBytes(sResult).Length;

		if (totalWidth - orgLen > 0) {
			sResult = new string(paddingChar, totalWidth - orgLen) + sResult;
		}

		return sResult;
	}
	#endregion

	#region 字串靠左對齊 +static string PadRightCHT(this string str, int totalWidth, char paddingChar)
	/// <summary>
	/// 字串靠左對齊，以指定的字元在右側補足長度，超過則截字(中文算2碼)。
	/// </summary>
	/// <param name="totalWidth">長度</param>
	/// <param name="paddingChar">替代字元</param>
	/// <returns></returns>
	public static string PadRightCHT(this string str, int totalWidth, char paddingChar) {
		string sResult = str.CutStr(totalWidth);
		int orgLen = Encoding.GetEncoding("big5").GetBytes(sResult).Length;

		if (totalWidth - orgLen > 0) {
			sResult = sResult + new string(paddingChar, totalWidth - orgLen);
		}

		return sResult;
	}
	#endregion
}