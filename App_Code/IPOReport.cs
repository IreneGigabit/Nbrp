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
	//private string _rectitleName = "";
	private DataTable _dtDmp = null;
	private DataTable _dtApcust = null;
	private DataTable _dtAgt = null;
	private DataTable _dtAnt = null;
	private DataTable _dtPrior = null;

    /// <summary>
    /// 報表代碼
    /// </summary>
    public string ReportCode { get; set; }

	/// <summary>
	/// 組合後的本所編號
	/// </summary>
	public string Seq {
		get { return _seq; }
		protected set { _seq = value; }
	}

	/// <summary>
	/// 收據抬頭種類→A:案件申請人/B:空白/C:案件申請人(代繳人)
	/// </summary>
	public string RectitleTitle { get; set; }

	/// <summary>
	/// 收據抬頭名稱
	/// </summary>
	public string RectitleName { get; set; }

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

    public IPOReport(string connStr, string in_scode, string in_no, string branch) : this(connStr, in_scode, in_no, branch, "") { }

	public IPOReport(string connStr, string in_scode, string in_no, string branch, string rectitle) {
		this._connStr = connStr;
		this._in_no = in_no;
		this._in_scode = in_scode;
		this._branch = branch;
		this.RectitleTitle = rectitle;
		this._conn = new DBHelper(connStr, false).Debug(true);

		this._dtDmp = new DataTable();
		_conn.DataTable("select *,''s_case1nm from vdmpall where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'", _dtDmp);//抓案件資料

		//專利類別
		string sCase1nm = "";
		switch (_dtDmp.Rows[0]["CASE1"].ToString().Substring(0, 2)) {
			case "IG":
				sCase1nm = "發明";
				break;
			case "UG":
				sCase1nm = "新型";
				break;
			case "DG":
				sCase1nm = "設計";
				break;
		}
		_dtDmp.Rows[0]["s_case1nm"] = sCase1nm;

		//若傳進的值是空的,以casp_dmp為主
		if (this.RectitleTitle == "") {
			this.RectitleTitle = _dtDmp.Rows[0]["receipt_title"].ToString();
		}
	}

    public IPOReport Init() {
		SetSeq();//組案件編號
		SetRectitleName();//抓收據抬頭
		SetApcust();//抓申請人
		SetAgent();//抓代理人
		SetAnt();//抓發明人/新型創作/設計人
		SetPrior();//抓優先權

        return this;
    }

	#region 關閉 +void Close()
	/// <summary>
	/// 關閉
	/// </summary>
	public void Close() {
		if (_conn != null) _conn.Dispose();
		this.Dispose();
	}
	#endregion

	#region 組本所編號 -void SetSeq()
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

	#region 抓收據抬頭 -void SetRectitleName()
	/// <summary>
	/// 抓收據抬頭
	/// </summary>
    private void SetRectitleName() {
        string RectitleNameStr = "";

        string kind = "'A'";
        if (this.ReportCode == "FC1") kind = "'C1','C2'";//申請權讓與
        if (this.ReportCode == "FC2") kind = "'C1','C2'";//專利權讓與
        if (this.ReportCode == "FL1_1") kind = "'D1','D2'";//專利權授權
        string SQL = "select b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename" +
                    " from dmp_apcust a " +
                    " inner join apcust b on a.apsqlno=b.apsqlno " +
                    " where a.kind in(" + kind + ") " +
                    " and in_scode='" + _in_scode + "' and in_no='" + _in_no + "' ";

        string Cname_string = "";
        using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
            while (dr.Read()) {
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

                if (RectitleNameStr != "") RectitleNameStr += "、";
                RectitleNameStr += Cname_string;
            }
        }

        if (this.RectitleTitle == "A") {//專利權人
            this.RectitleName = RectitleNameStr;
        } else if (this.RectitleTitle == "C") {//專利權人(代繳人)
            this.RectitleName = RectitleNameStr + "(代繳人：聖島國際專利商標聯合事務所)";
        } else {//空白
            this.RectitleName = "";
        }

    }
	#endregion

	#region 抓申請人 -void SetApcust()
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

	#region 抓代理人 -void SetAgent()
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

	#region 抓發明人/新型創作/設計人資料 -void SetAnt()
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

    #region 抓相對人 +DataTable GetApAnt()
    /// <summary>
    /// 抓相對人
    /// </summary>
    /// <param name="kind">
    /// <para>A=發明/創作人</para>
    /// <para>C1=讓與人 C2=受讓人</para>
    /// <para>D1=授權人 D2=被授權人</para>
    /// <para>E=異議/舉發人</para>
    /// <para>X=相關對照人</para>
    /// <para>F1=質權人 F2=出質人</para>
    /// <para>G1=繼承人 G2=被繼承人</para>
    /// <para>H1=受託人 H2=委託人</para>
    /// </param>
    public DataTable GetApAnt(string kind) {
        string SQL = "select a.kind,a.seqno,a.apsqlno,a.ap_cname1,a.ap_cname2,b.apcust_no,b.apclass " +
        ",b.ap_ename1,b.ap_ename2,b.ap_crep,b.ap_erep " +
        ",b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename " +
        ",b.ap_country,b.ap_zip,b.ap_addr1,b.ap_addr2 " +
        ",b.ap_eaddr1,b.ap_eaddr2,b.ap_eaddr3,b.ap_eaddr4 " +
        ",''apclass_name,''Country_name,''Title_cname,''Cname_string,''Title_ename,''Ename_string " +
        " from dmp_apcust a " +
        " inner join apcust b on a.apsqlno=b.apsqlno " +
        " where a.in_scode='" + _in_scode + "' and a.in_no='" + _in_no + "' " +
        " and kind='" + kind + "' " +
        " union " +
        " select a.kind,a.seqno,0 as apsqlno,a.ant_cname1 as ap_cname1,a.ant_cname2 as ap_cname2,a.ant_id as apcust_no " +
        ",(case len(a.ant_id) when 10 then 'B' else '' end) as apclass " +
        ",a.ant_ename1 as ap_ename1,a.ant_ename2 as ap_ename2,a.ant_crep as ap_crep,a.ant_erep as ap_erep " +
        ",a.ant_fcname as ap_fcname,a.ant_lcname as ap_lcname,a.ant_fename as ap_fename,a.ant_lename as ap_lename " +
        ",a.ant_country as ap_country,a.ant_zip as ap_zip,a.ant_addr1 as ap_addr1,a.ant_addr2 as ap_addr2 " +
        ",'' as ap_eaddr1,'' as ap_eaddr2,'' as ap_eaddr3,'' as ap_eaddr4 " +
        ",''apclass_name,''Country_name,''Title_cname,''Cname_string,''Title_ename,''Ename_string " +
        " from dmp_ant a " +
        " where a.in_scode='" + _in_scode + "' and a.in_no='" + _in_no + "' " +
        " and kind='" + kind + "' " +
        " order by a.seqno ";

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
            else if (dt.Rows[i]["apclass"].ToString().Left(1) == "B") {
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
            else if (dt.Rows[i]["apclass"].ToString().Left(1) == "C" || dt.Rows[i]["apclass"].ToString().Left(1) != "B") {
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
        return dt;
    }
    #endregion

	#region 抓優先權資料 -void SetPrior()
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

	#region 產生基本資料表+void AppendBaseData(string baseDocName, string antTitle)
	/// <summary>
	/// 產生基本資料表
	/// </summary>
	public void AppendBaseData(string baseDocName, string antTitle) {
		CopyBlock(baseDocName, "base_title");
		//申請人
		for (int i = 0; i < Apcust.Rows.Count; i++) {
			CopyBlock(baseDocName, "base_apcust");
            ReplaceBookmark("base_ap_type", "申請人");
            ReplaceBookmark("base_ap_num", (i + 1).ToString());
			ReplaceBookmark("base_ap_country", Apcust.Rows[i]["Country_name"].ToString());
			ReplaceBookmark("ap_class", Apcust.Rows[i]["apclass_name"].ToString());
			if (Apcust.Rows[i]["ap_country"].ToString() == "T") {
				ReplaceBookmark("apcust_no", Apcust.Rows[i]["apcust_no"].ToString());
			} else {
				ReplaceBookmark("apcust_no", "", true);
			}
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
		for (int i = 0; i < Agent.Rows.Count; i++) {
			CopyBlock(baseDocName, "base_agent");
            ReplaceBookmark("agt_type1", "");
			ReplaceBookmark("agt_idno1", Agent.Rows[i]["agt_idno1"].ToString());
			ReplaceBookmark("agt_id1", Agent.Rows[i]["agt_id1"].ToString());
			ReplaceBookmark("base_agt_name1", Agent.Rows[i]["agt_name1"].ToString());
			ReplaceBookmark("agt_zip1", Agent.Rows[i]["agt_zip"].ToString());
			ReplaceBookmark("agt_addr1", Agent.Rows[i]["agt_addr"].ToString());
			ReplaceBookmark("agatt_tel1", Agent.Rows[i]["agt_tel"].ToString());
			ReplaceBookmark("agatt_fax1", Agent.Rows[i]["agt_fax"].ToString());

            ReplaceBookmark("agt_type2", "");
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
				CopyBlock(baseDocName, "base_ant");
				ReplaceBookmark("base_ant_num", antTitle + (i + 1).ToString());
				ReplaceBookmark("base_ant_country", Ant.Rows[i]["Country_name"].ToString());
				if (Ant.Rows[i]["ant_country"].ToString() == "T") {
					ReplaceBookmark("ant_id", Ant.Rows[i]["ant_id"].ToString());
				} else {
					ReplaceBookmark("ant_id", "", true);
				}
				ReplaceBookmark("base_ant_cname", Ant.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("base_ant_ename", Ant.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
				AddParagraph();
			}
		}

		CopyPageFoot(baseDocName, false);//頁尾
	}
	#endregion

	#region 更新列印狀態 +void SetPrint()
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