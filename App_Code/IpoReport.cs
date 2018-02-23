using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Web;

/// <summary>
/// 產生智慧局電子申請書用
/// </summary>
public class IPOReport : OpenXmlHelper {
	private string _connStr = null;
	private string _in_no = "";
	private string _in_scode = "";
	private string _branch = "";
	private DBHelper _conn = null;

	private string _seq = "";
	private string _rectitleName = "";
	private DataTable _dtDmp = null;
	private DataTable _dtApcust = null;
	private DataTable _dtAgt = null;
	private DataTable _dtAnt = null;
	private DataTable _dtPrior = null;

	/// <summary>
	/// 組合後的本所編號
	/// </summary>
	public string Seq {
		get { return _seq; }
		protected set { _seq = value; }
	}

	/// <summary>
	/// 收據種類
	/// </summary>
	public string RectitleTitle { get; set; }

	/// <summary>
	/// 收據抬頭
	/// </summary>
	public string RectitleName {
		get { return _rectitleName; }
		protected set { _rectitleName = value; }
	}

	/// <summary>
	/// 案件資料
	/// </summary>
	public DataTable Dmp {
		get { return _dtDmp; }
		protected set { _dtDmp = value; }
	}

	/// <summary>
	/// 申請人資料
	/// </summary>
	public DataTable Apcust {
		get { return _dtApcust; }
		protected set { _dtApcust = value; }
	}

	/// <summary>
	/// 代理人資料
	/// </summary>
	public DataTable Agent {
		get { return _dtAgt; }
		protected set { _dtAgt = value; }
	}

	/// <summary>
	/// 發明人/新型創作/設計人資料
	/// </summary>
	public DataTable Ant {
		get { return _dtAnt; }
		protected set { _dtAnt = value; }
	}

	/// <summary>
	/// 優先權資料
	/// </summary>
	public DataTable Prior {
		get { return _dtPrior; }
		protected set { _dtPrior = value; }
	}

	public IPOReport() {

	}

	public IPOReport(string connStr, string in_scode, string in_no, string branch) {
		this._connStr = connStr;
		this._in_no = in_no;
		this._in_scode = in_scode;
		this._branch = branch;
		this._conn = new DBHelper(connStr, false).Debug(false);

		this._dtDmp = new DataTable();
		_conn.DataTable("select * from vdmpall where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'", _dtDmp);//抓案件資料

		SetSeq();//組案件編號
		SetRectitleName();//抓收據抬頭
		SetApcust();//抓申請人
		SetAgent();//抓代理人
		SetAnt();//抓發明人/新型創作/設計人
		SetPrior();//抓優先權
	}

	//電子收據第2階段上線後要廢除RectitleTitle參數
	public IPOReport(string connStr, string in_scode, string in_no, string branch, string rectitle) {
		this._connStr = connStr;
		this._in_no = in_no;
		this._in_scode = in_scode;
		this._branch = branch;
		this.RectitleTitle = rectitle;
		this._conn = new DBHelper(connStr, false).Debug(false);

		this._dtDmp = new DataTable();
		_conn.DataTable("select * from vdmpall where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'", _dtDmp);//抓案件資料

		SetSeq();//組案件編號
		SetRectitleName();//抓收據抬頭
		SetApcust();//抓申請人
		SetAgent();//抓代理人
		SetAnt();//抓發明人/新型創作/設計人
		SetPrior();//抓優先權
	}

	#region 關閉
	/// <summary>
	/// 關閉
	/// </summary>
	public void Close() {
		if (_conn != null) _conn.Dispose();
		this.Dispose();
	}
	#endregion

	#region 組本所編號
	/// <summary>
	/// 組本所編號
	/// </summary>
	private void SetSeq() {
		string lseq = _branch + "P" + _dtDmp.Rows[0]["seq"];
		if (_dtDmp.Rows[0]["seq1"].ToString() != "_") {
			lseq += "-" + _dtDmp.Rows[0]["seq1"];
		}
		this.Seq = lseq;
	}
	#endregion

	#region 抓收據抬頭
	/// <summary>
	/// 抓收據抬頭
	/// </summary>
	private void SetRectitleName() {
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

			if (this.RectitleTitle == "A") {//專利權人
				RectitleName = Cname_string;
			} else if (this.RectitleTitle == "C") {//專利權人(代繳人)
				RectitleName = Cname_string + "(代繳人：聖島國際專利商標聯合事務所)";
			} else {//空白
				RectitleName = "";
			}

			/*
			string receipt_title = _dtDmp.Rows[0]["receipt_title"].ToString();
			if (receipt_title == "A") {//專利權人
				RectitleName = Cname_string;
			} else if (receipt_title == "C") {//專利權人(代繳人)
				RectitleName = Cname_string + "(代繳人：聖島國際專利商標聯合事務所)";
			} else if (receipt_title == "B") {//空白
				RectitleName = "";
			}*/
		}

		this.RectitleName = RectitleName;
	}
	#endregion

	#region 抓申請人
	/// <summary>
	/// 抓申請人
	/// </summary>
	private void SetApcust() {
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
					dt.Rows[i]["apclass_name"] = "法人公司機關學校/商號行號工廠";
					dt.Rows[i]["Title_cname"] = "中文名稱";
					dt.Rows[i]["Title_ename"] = "英文名稱";
					dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				} else {
					dt.Rows[i]["apclass_name"] = "自然人";
					dt.Rows[i]["Title_cname"] = "中文姓名";
					dt.Rows[i]["Title_ename"] = "英文姓名";
				}

				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_fename"].ToString().Trim() + "," + dt.Rows[i]["ap_lename"].ToString().Trim();
				if (dt.Rows[i]["Ename_string"].ToString() == ",") {
					dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
				}
			}
		}
		this.Apcust = dt;
	}
	#endregion

	#region 抓代理人
	/// <summary>
	/// 抓代理人
	/// </summary>
	private void SetAgent() {
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
		this.Agent = dt;
	}
	#endregion

	#region 抓發明人/新型創作/設計人資料
	/// <summary>
	/// 抓發明人/新型創作/設計人資料
	/// </summary>
	private void SetAnt() {
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
		this.Ant = dt;
	}
	#endregion

	#region 抓優先權資料
	/// <summary>
	/// 抓優先權資料
	/// </summary>
	private void SetPrior() {
		string SQL = "SELECT a.prior_yn, a.prior_no, a.prior_country, a.prior_date, a.mprior_access, a.prior_case1 " +
			", c.coun_code, c.coun_cname, c.coun_ename " +
			", (SELECT mark1 FROM cust_code WHERE code_type = 'case1' AND cust_code = a.prior_case1) AS case1nm_T " +
			", (SELECT code_name FROM cust_code WHERE code_type = 'pecase1' AND cust_code = a.prior_case1) AS case1nm " +
			", isnull(c.coun_code,'')+isnull(c.coun_cname,'') Country_name " +
			" FROM dmp_prior AS a " +
			" INNER JOIN vdmpall AS b ON a.seq = b.seq AND a.seq1 = b.seq1 " +
			" LEFT JOIN sysctrl.dbo.IPO_country AS c ON a.prior_country = c.ref_coun_code " +
			" WHERE b.in_scode = '" + _in_scode + "' " +
			" AND b.in_no = '" + _in_no + "' " +
			" AND a.prior_yn = 'Y'";

		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);
		this.Prior = dt;
	}
	#endregion

	#region 產生基本資料表
	/// <summary>
	/// 產生基本資料表
	/// </summary>
	public void AppendBaseData(string baseDocName, string antTitle) {
		CopyBlock(baseDocName, "base_title");
		//申請人
		for (int i = 0; i < Apcust.Rows.Count; i++) {
			CopyBlock(baseDocName, "base_apcust1");
			ReplaceBookmark("base_ap_num", (i + 1).ToString());
			ReplaceBookmark("base_ap_country", Apcust.Rows[i]["Country_name"].ToString());
			ReplaceBookmark("ap_class", Apcust.Rows[i]["apclass_name"].ToString());
			if (Apcust.Rows[i]["ap_country"].ToString() == "T") {
				CopyBlock(baseDocName, "base_apcust2");
				ReplaceBookmark("apcust_no", Apcust.Rows[i]["apcust_no"].ToString());
			}
			CopyBlock(baseDocName, "base_apcust3");
			ReplaceBookmark("base_ap_cname_title", Apcust.Rows[i]["Title_cname"].ToString());
			ReplaceBookmark("base_ap_ename_title", Apcust.Rows[i]["Title_ename"].ToString());
			ReplaceBookmark("base_ap_cname", Apcust.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
			ReplaceBookmark("base_ap_ename", Apcust.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
			ReplaceBookmark("ap_live_country", Apcust.Rows[i]["Country_name"].ToString());
			ReplaceBookmark("ap_zip", Apcust.Rows[i]["ap_zip"].ToString());
			string ap_addr = Apcust.Rows[i]["ap_addr1"].ToString().ToXmlUnicode()
				+ Apcust.Rows[i]["ap_addr2"].ToString().ToXmlUnicode();
			ReplaceBookmark("ap_addr", ap_addr);
			string ap_eddr = Apcust.Rows[i]["ap_eaddr1"].ToString().ToXmlUnicode(true)
				+ Apcust.Rows[i]["ap_eaddr2"].ToString().ToXmlUnicode(true)
				+ Apcust.Rows[i]["ap_eaddr3"].ToString().ToXmlUnicode(true)
				+ Apcust.Rows[i]["ap_eaddr4"].ToString().ToXmlUnicode(true);
			ReplaceBookmark("ap_eddr", ap_eddr);
			ReplaceBookmark("ap_crep", Apcust.Rows[i]["ap_crep"].ToString().ToXmlUnicode());
			ReplaceBookmark("ap_erep", Apcust.Rows[i]["ap_erep"].ToString().ToXmlUnicode(true));
		}
		//代理人
		CopyBlock(baseDocName, "base_agent");
		for (int i = 0; i < Agent.Rows.Count; i++) {
			CopyBlock(baseDocName, "base_apcust");
			ReplaceBookmark("agt_idno1", Agent.Rows[i]["agt_idno1"].ToString());
			ReplaceBookmark("agt_id1", Agent.Rows[i]["agt_id1"].ToString());
			ReplaceBookmark("base_agt_name1", Agent.Rows[i]["agt_name1"].ToString());
			ReplaceBookmark("agt_zip1", Agent.Rows[i]["agt_zip"].ToString());
			ReplaceBookmark("agt_addr1", Agent.Rows[i]["agt_addr"].ToString());
			ReplaceBookmark("agatt_tel1", Agent.Rows[i]["agt_tel"].ToString());
			ReplaceBookmark("agatt_fax1", Agent.Rows[i]["agt_fax"].ToString());
			ReplaceBookmark("agt_idno2", Agent.Rows[i]["agt_idno2"].ToString());
			ReplaceBookmark("agt_id2", Agent.Rows[i]["agt_id2"].ToString());
			ReplaceBookmark("base_agt_name2", Agent.Rows[i]["agt_name2"].ToString());
			ReplaceBookmark("agt_zip2", Agent.Rows[i]["agt_zip"].ToString());
			ReplaceBookmark("agt_addr2", Agent.Rows[i]["agt_addr"].ToString());
			ReplaceBookmark("agatt_tel2", Agent.Rows[i]["agt_tel"].ToString());
			ReplaceBookmark("agatt_fax2", Agent.Rows[i]["agt_fax"].ToString());
		}

		if (antTitle != "") {
			//發明人/新型創作/設計人
			for (int i = 0; i < Ant.Rows.Count; i++) {
				CopyBlock(baseDocName, "base_ant1");
				ReplaceBookmark("base_ant_num", antTitle + (i + 1).ToString());
				ReplaceBookmark("base_ant_country", Ant.Rows[i]["Country_name"].ToString());
				if (Ant.Rows[i]["ant_country"].ToString() == "T") {
					CopyBlock(baseDocName, "base_ant2");
					ReplaceBookmark("ant_id", Ant.Rows[i]["ant_id"].ToString());
				}
				CopyBlock(baseDocName, "base_ant3");
				ReplaceBookmark("base_ant_cname", Ant.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("base_ant_ename", Ant.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
				AddParagraph();
			}
		}

		CopyPageFoot(baseDocName, false);//頁尾
	}
	#endregion

	#region 更新列印狀態
	/// <summary>
	/// 更新列印狀態
	/// </summary>
	public void SetPrint() {
		string SQL = "update case_dmp set new='P'+substring(NEW,2,50) " +
					",receipt_title='" + this.RectitleTitle + "' " +
					",rectitle_name='" + this.RectitleName + "' " +
					"where in_scode='" + this._in_scode + "' and in_no='" + this._in_no + "'";
		_conn.ExecuteNonQuery(SQL);
	}
	#endregion
}