using System;
using System.Collections.Generic;
using System.Web;
using System.Globalization;

public static class DateTimeExt {
	/// <summary>
	/// To the full taiwan date.
	/// </summary>
	/// <param name="datetime">The datetime.</param>
	/// <returns></returns>
	public static string ToLongTwDate(this DateTime datetime) {
		TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

		return string.Format("民國{0}年{1}月{2}日",
			taiwanCalendar.GetYear(datetime).ToString().PadLeft(3, '0'),
			datetime.Month.ToString().PadLeft(2, '0'),
			datetime.Day.ToString().PadLeft(2, '0'));
	}

	/// <summary>
	/// To the simple taiwan date.
	/// </summary>
	/// <param name="datetime">The datetime.</param>
	/// <returns></returns>
	public static string ToShortTwDate(this DateTime datetime) {
		TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

		return string.Format("{0}/{1}/{2}",
			taiwanCalendar.GetYear(datetime).ToString().PadLeft(3, '0'),
			datetime.Month.ToString().PadLeft(2, '0'),
			datetime.Day.ToString().PadLeft(2, '0'));
	}

	/// <summary>
	/// To the simple taiwan date.
	/// </summary>
	/// <param name="datetime">The datetime.</param>
	/// <returns></returns>
	public static int GetTwYear(this DateTime datetime) {
		TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

		return taiwanCalendar.GetYear(datetime);
	}

	/// <summary>
	/// 轉成西元年informix  99/99/9999(月/日/年) 給資料型態為SmallDateTime 使用
	/// </summary>
	/// <param name="string">The date string.</param>
	/// <returns></returns>
	public static string ToD9Date(this string dateString) {
		if (dateString == null || dateString.Trim() == "") {
			return "";
		} else {
			DateTime dt = Convert.ToDateTime(dateString);
			return string.Format("{0}/{1}/{2}"
				, dt.Month.ToString().PadLeft(2, '0')
				, dt.Day.ToString().PadLeft(2, '0')
				, dt.Year.ToString());
		}
	}

}