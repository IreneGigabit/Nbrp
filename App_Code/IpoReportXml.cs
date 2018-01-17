using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Web;

/// <summary>
/// 產生智慧局電子申請書用
/// </summary>
public class IpoReportXml {
	protected string _connStr = null;
	protected string _in_no = "";
	protected string _in_scode = "";
	protected string _branch = "";
	protected string _receipt_title = "";
	protected DBHelper _conn = null;
	protected DataTable dtDmp {get;set;}

	public IpoReportXml() {

	}

	public IpoReportXml(string connStr, string in_scode, string in_no, string branch, string receipt_title) {
		this._connStr = connStr;
		this._in_no = in_no;
		this._in_scode = in_scode;
		this._branch = branch;
		this._receipt_title = receipt_title;
		this._conn = new DBHelper(connStr, false).Debug(false);
		this.dtDmp = new DataTable();
		_conn.DataTable("select * from vdmpall where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'", dtDmp);
	}

	/// <summary>
	/// 關閉
	/// </summary>
	public IpoReportXml Close() {
		_conn.Dispose();
		return this;
	}

	#region 取得組合後的本所編號
	/// <summary>
	/// 取得組合後的本所編號
	/// </summary>
	public string getSeq() {
		string lseq = _branch + "P" + dtDmp.Rows[0]["seq"];
		if (dtDmp.Rows[0]["seq1"].ToString() != "_") {
			lseq += "-" + dtDmp.Rows[0]["seq1"];
		}
		return lseq;
	}
	#endregion

	#region 取得案件資料
	/// <summary>
	/// 取得案件資料
	/// </summary>
	public DataTable getDmp() {
		return dtDmp;
	}
	#endregion

	#region 取得申請人資料
	/// <summary>
	/// 取得申請人資料
	/// </summary>
	public DataTable GetBaseApcust() {
		string SQL = "select b.ap_zip,b.ap_crep,b.ap_erep,b.ap_eaddr1,b.ap_eaddr2,b.ap_eaddr3,b.ap_eaddr4,b.ap_addr1,b.ap_addr2 " +
		",b.apcust_no,b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename " +
		",''apclass_name,''Country_name,''Title_cname,''Cname_string,''Title_ename,''Ename_string " +
		" from dmp_apcust a " +
		" inner join apcust b on a.apsqlno=b.apsqlno " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'  and kind='A'";
		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);

		for (int i = 0; i < dt.Rows.Count; i++) {
			SQL = " select  isnull(b.coun_code,'')+isnull(b.coun_cname,'') Country_name " +
				"From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code " +
			 " where a.coun_code = '" + dt.Rows[i]["ap_country"] + "'";
			dt.Rows[i]["Country_name"] = (_conn.ExecuteScalar(SQL) ?? "").ToString();

			//本國公司
			if (dt.Rows[i]["apclass"].ToString().Left(1) == "A") {
				if (dt.Rows[i]["apclass"].ToString() == "AD") {
					dt.Rows[i]["apclass_name"] = "商號行號工廠";
				} else {
					dt.Rows[i]["apclass_name"] = "法人公司機關學校";
				}
				dt.Rows[i]["Title_cname"] = "中文名稱";
				dt.Rows[i]["Title_ename"] = "英文名稱";
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
			}
			//本國自然人
			if (dt.Rows[i]["apclass"].ToString().Left(1) == "B") {
				dt.Rows[i]["apclass_name"] = "自然人";
				dt.Rows[i]["Title_cname"] = "中文姓名";
				dt.Rows[i]["Title_ename"] = "英文姓名";
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_fcname"].ToString().Trim() + "," + dt.Rows[i]["ap_lcname"].ToString().Trim();
				if (dt.Rows[i]["Cname_string"].ToString() == ",") {
					dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				}

				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_fename"].ToString().Trim() + "," + dt.Rows[i]["ap_lename"].ToString().Trim();
				if (dt.Rows[i]["Ename_string"].ToString() == ",") {
					dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
				}
			}

			//20161206外國人/公司增加判斷-若有填寫申請人姓&名,則顯示姓名
			if (dt.Rows[i]["apclass"].ToString().Left(1) == "C") {
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_fcname"].ToString().Trim() + "," + dt.Rows[i]["ap_lcname"].ToString().Trim();
				if (dt.Rows[i]["Cname_string"].ToString() == ",") {
					dt.Rows[i]["Title_cname"] = "中文名稱";
					dt.Rows[i]["Title_ename"] = "英文名稱";
					dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				} else {
					dt.Rows[i]["Title_cname"] = "中文姓名";
					dt.Rows[i]["Title_ename"] = "英文姓名";
				}

				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_fename"].ToString().Trim() + "," + dt.Rows[i]["ap_lename"].ToString().Trim();
				if (dt.Rows[i]["Ename_string"].ToString() == ",") {
					dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
				}
			}
		}
		return dt;
	}
	#endregion

	#region 取得代理人資料
	/// <summary>
	/// 取得代理人資料
	/// </summary>
	public DataTable GetBaseAgt() {
		string SQL = " Select b.agt_fax,b.agt_tel,b.agt_addr,b.agt_zip,b.agt_id1,b.agt_id2,b.agt_idno1,b.agt_idno2,b.agt_name1,b.agt_name2 from dmp a " +
		" inner join vdmpall c on a.dmp_sqlno = c.dmp_sqlno " +
		" inner join agt b on c.nagt_no = b.agt_no " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'";
		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);

		for (int i = 0; i < dt.Rows.Count; i++) {
			dt.Rows[i]["agt_idno1"] = dt.Rows[i]["agt_idno1"].ToString().Trim().PadLeft(5, '0');
			dt.Rows[i]["agt_idno2"] = dt.Rows[i]["agt_idno2"].ToString().Trim().PadLeft(5, '0');
			dt.Rows[i]["agt_name1"] = dt.Rows[i]["agt_name1"].ToString().Trim().Left(1) + "," + dt.Rows[i]["agt_name1"].ToString().Trim().Substring(1);
			dt.Rows[i]["agt_name2"] = dt.Rows[i]["agt_name2"].ToString().Trim().Left(1) + "," + dt.Rows[i]["agt_name2"].ToString().Trim().Substring(1);
		}
		return dt;
	}
	#endregion

	#region 取得明人/新型創作/設計人
	/// <summary>
	/// 取得發明人/新型創作/設計人
	/// </summary>
	public DataTable GetBaseAnt() {
		string SQL = " Select ant_id,ant_country,ant_cname1,ant_cname2,ant_ename1,ant_ename2,ant_fcname,ant_lcname,ant_fename,ant_lename " +
			",''Country_name,''Cname_string,''Ename_string " +
			"from dmp_ant " +
			" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'";
		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);

		for (int i = 0; i < dt.Rows.Count; i++) {
			SQL = " select  isnull(b.coun_code,'')+isnull(b.coun_cname,'') Country_name " +
				"From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code " +
			 " where a.coun_code = '" + dt.Rows[i]["ant_country"] + "'";
			dt.Rows[i]["Country_name"] = (_conn.ExecuteScalar(SQL) ?? "").ToString();

			dt.Rows[i]["Cname_string"] = dt.Rows[i]["ant_fcname"].ToString().Trim() + "," + dt.Rows[i]["ant_lcname"].ToString().Trim();
			if (dt.Rows[i]["Cname_string"].ToString() == ",") {
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ant_cname1"].ToString().Trim() + dt.Rows[i]["ant_cname2"].ToString().Trim();
			}

			dt.Rows[i]["Ename_string"] = dt.Rows[i]["ant_fename"].ToString().Trim() + "," + dt.Rows[i]["ant_lename"].ToString().Trim();
			if (dt.Rows[i]["Ename_string"].ToString() == ",") {
				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ant_ename1"].ToString().Trim() + dt.Rows[i]["ant_ename2"].ToString().Trim();
			}
		}
		return dt;
	}

	#endregion

	#region 產生申請人區塊
	/// <summary>
	/// 產生申請人區塊
	/// </summary>
	public string GetApcustBlock(string xml) {
		string returnStr = "";

		string SQL = "select b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename " +
		" from dmp_apcust a " +
		" inner join apcust b on a.apsqlno=b.apsqlno " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "' and kind='A' ";

		int i = 0;
		using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
			while (dr.Read()) {
				i++;
				returnStr += xml;
				//抬頭
				returnStr = returnStr.Replace("#apply_num#", i.ToString());
				//國籍
				SQL = " select isnull(b.coun_code,'')+isnull(b.coun_cname,'') Country_name " +
				"From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code " +
				" where a.coun_code = '" + dr.GetString("ap_country") + "'";
				using (DBHelper _conn1 = new DBHelper(_connStr)) {
					string Country_name = (_conn1.ExecuteScalar(SQL) ?? "").ToString();
					returnStr = returnStr.Replace("#ap_country#", Country_name);
				}

				string Title_cname = "";
				string Title_ename = "";
				string Cname_string = "";
				string Ename_string = "";

				//本國公司
				if (dr.GetString("apclass").Left(1) == "A") {
					Title_cname = "中文名稱";
					Title_ename = "英文名稱";
					Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					Ename_string = dr.GetString("ap_ename1") + dr.GetString("ap_ename2");
				}

				//本國自然人
				if (dr.GetString("apclass").Left(1) == "B") {
					Title_cname = "中文姓名";
					Title_ename = "英文姓名";
					Cname_string = dr.GetString("ap_fcname") + "," + dr.GetString("ap_lcname");
					if (Cname_string == ",") {
						Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					}

					Ename_string = dr.GetString("ap_fename") + "," + dr.GetString("ap_lename");
					if (Ename_string == ",") {
						Ename_string = dr.GetString("ap_ename1") + dr.GetString("ap_ename2");
					}
				}

				//20161206外國人/公司增加判斷-若有填寫申請人姓&名,則顯示姓名
				if (dr.GetString("apclass").Left(1) == "C") {
					Cname_string = dr.GetString("ap_fcname") + "," + dr.GetString("ap_lcname");
					if (Cname_string == ",") {
						Title_cname = "中文名稱";
						Title_ename = "英文名稱";
						Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					} else {
						Title_cname = "中文姓名";
						Title_ename = "英文姓名";
					}

					Ename_string = dr.GetString("ap_fename") + "," + dr.GetString("ap_lename");
					if (Ename_string == ",") {
						Ename_string = dr.GetString("ap_ename1") + dr.GetString("ap_ename2");
					}
				}

				returnStr = returnStr.Replace("#ap_cname1_title#", Title_cname);
				returnStr = returnStr.Replace("#ap_ename1_title#", Title_ename);
				returnStr = returnStr.Replace("#ap_cname1#", Cname_string.ToXmlUnicode());
				returnStr = returnStr.Replace("#ap_ename1#", Ename_string.ToXmlUnicode());
			}
		}
		return returnStr;
	}
	#endregion

	#region 產生代理人1 & 代理人2 區塊
	/// <summary>
	/// 產生代理人1 & 代理人2 區塊
	/// </summary>
	public string GetAgtBlock(string xml) {
		string returnStr = "";

		string SQL = " Select b.agt_name1,b.agt_name2 from dmp a " +
		" inner join vdmpall c on a.dmp_sqlno = c.dmp_sqlno " +
		" inner join agt b on c.nagt_no = b.agt_no " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'";

		using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
			if (dr.Read()) {
				//代理人1
				returnStr += xml;
				returnStr = returnStr.Replace("#agt_num#", "1");
				returnStr = returnStr.Replace("#agt_name#", dr.GetString("agt_name1").Left(1) + "," + dr.GetString("agt_name1").Substring(1));

				//代理人2
				returnStr += xml;
				returnStr = returnStr.Replace("#agt_num#", "2");
				returnStr = returnStr.Replace("#agt_name#", dr.GetString("agt_name2").Left(1) + "," + dr.GetString("agt_name2").Substring(1));
			}
		}
		return returnStr;
	}
	#endregion

	#region 產生發明人/新型創作/設計人區塊
	/// <summary>
	/// 產生發明人/新型創作/設計人區塊
	/// </summary>
	public string GetAntBlock(string xml, string type_str) {
		string returnStr = "";

		string SQL = " Select ant_country,ant_cname1,ant_cname2,ant_ename1,ant_ename2,ant_fcname,ant_lcname,ant_fename,ant_lename from dmp_ant " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'";

		int i = 0;
		using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
			while (dr.Read()) {
				i++;
				returnStr += xml;
				//抬頭
				returnStr = returnStr.Replace("#ant_num#", type_str + i.ToString());
				//國籍
				SQL = " select isnull(b.coun_code,'')+isnull(b.coun_cname,'') Country_name " +
				"From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code " +
				" where a.coun_code = '" + dr.GetString("ant_country") + "'";
				using (DBHelper _conn1 = new DBHelper(_connStr)) {
					string Country_name = (_conn1.ExecuteScalar(SQL) ?? "").ToString();
					returnStr = returnStr.Replace("#ant_country#", Country_name);
				}

				string Cname_string = "";
				string Ename_string = "";

				Cname_string = dr.GetString("ant_fcname") + "," + dr.GetString("ant_lcname");
				if (Cname_string == ",") {
					Cname_string = dr.GetString("ant_cname1") + dr.GetString("ant_cname2");
				}

				Ename_string = dr.GetString("ant_fename") + "," + dr.GetString("ant_lename");
				if (Cname_string == ",") {
					Cname_string = dr.GetString("ant_ename1") + dr.GetString("ant_ename2");
				}

				returnStr = returnStr.Replace("#ant_cname#", Cname_string.ToXmlUnicode());
				returnStr = returnStr.Replace("#ant_ename#", Ename_string.ToXmlUnicode());
			}
		}

		return returnStr;
	}

	#endregion

	#region 產生主張優先權區塊
	/// <summary>
	/// 產生主張優先權區塊
	/// </summary>
	public string GetPriorBlock(string xml1, string xml_JA, string xml_KO,string xml_space) {
		string returnStr = "";

		string SQL = "SELECT a.prior_yn, a.prior_no, a.prior_country, a.prior_date, a.mprior_access, a.prior_case1 " +
			", c.coun_code, c.coun_cname, c.coun_ename" +
			", (SELECT mark1 FROM cust_code WHERE code_type = 'case1' AND cust_code = a.prior_case1) AS case1nm_T " +
			", (SELECT code_name FROM cust_code WHERE code_type = 'pecase1' AND cust_code = a.prior_case1) AS case1nm " +
			" FROM dmp_prior AS a " +
			" INNER JOIN vdmpall AS b ON a.seq = b.seq AND a.seq1 = b.seq1 " +
			" LEFT JOIN sysctrl.dbo.IPO_country AS c ON a.prior_country = c.ref_coun_code " +
			" WHERE b.in_scode = '" + _in_scode + "' " +
			" AND b.in_no = '" + _in_no + "' " +
			" AND a.prior_yn = 'Y' ";

		int i = 0;
		using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
			while (dr.Read()) {
				i++;
				returnStr += xml1;
				returnStr = returnStr.Replace("#prior_num#", i.ToString());
				returnStr = returnStr.Replace("#prior_country#", dr.GetString("coun_code") + dr.GetString("coun_cname"));
				returnStr = returnStr.Replace("#prior_date#", (dr.GetNullDateTime("prior_date") != null ? dr.GetNullDateTime("prior_date").Value.ToString("yyyy/MM/dd") : ""));
				returnStr = returnStr.Replace("#prior_no#", dr.GetString("prior_no"));

				switch (dr.GetString("prior_country")) {
					case "JA":
						returnStr += xml_JA;
						returnStr = returnStr.Replace("#case1nm#", dr.GetString("case1nm"));
						returnStr = returnStr.Replace("#mprior_access#", dr.GetString("mprior_access"));
						break;
					case "KO":
						returnStr += xml_KO;
						returnStr = returnStr.Replace("#mprior_access#", "交換");
						break;
				}
			}
		}
		return returnStr;
	}
	#endregion

	#region 取得收據抬頭
	/// <summary>
	/// 取得收據抬頭
	/// </summary>
	public string getRectitleName(string receipt_title) {
		string RectitleName = "";

		//申請人(只抓一個)
		string SQL = "select b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename" +
					" from dmp_apcust a,apcust b " +
					" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'  and kind='A' and a.apsqlno=b.apsqlno";

		string Cname_string = "";
		using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
			if (dr.Read()) {
				//本國公司
				if (dr.GetString("apclass").Left(1) == "A") {
					Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
				}

				//本國自然人
				if (dr.GetString("apclass").Left(1) == "B") {
					Cname_string = dr.GetString("ap_fcname") + "" + dr.GetString("ap_lcname");
					if (Cname_string == "") {
						Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					}
				}

				//20161206外國人/公司增加判斷-若有填寫申請人姓&名,則顯示姓名
				if (dr.GetString("apclass").Left(1) == "C") {
					Cname_string = dr.GetString("ap_fcname") + "" + dr.GetString("ap_lcname");
					if (Cname_string == "") {
						Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					}
				}
			}

			if (receipt_title == "A") {//專利權人
				RectitleName = Cname_string;
			} else if (receipt_title == "C") {//專利權人(代繳人)
				RectitleName = Cname_string + "(代繳人：聖島國際專利商標聯合事務所)";
			} else if (receipt_title == "B") {//空白
				RectitleName = "";
			}
		}

		return RectitleName;
	}
	#endregion
}