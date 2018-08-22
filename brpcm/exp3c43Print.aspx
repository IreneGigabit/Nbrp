<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Odbc" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>

<script runat="server">
    protected OpenXmlHelper Rpt = new OpenXmlHelper();
    int right=0;
    string se_scode, se_branch, se_branchnm;
    string se_dept, se_syscode, se_seqtitle;
    string qsort, hend, qcust_area, qcust_seq;
    string qscase_date, qecase_date, qseq, qseq1;
    string qcust_prod, qcustprod_no, qcase1, qscode1;
        
    private void Page_Load(System.Object sender, System.EventArgs e) {
        Response.CacheControl = "Private";
        Response.AddHeader("Pragma", "no-cache");
        Response.Expires = -1;
        Response.Clear();

        right=Convert.ToInt32(Request["right"] ?? "0");
        se_scode = (Request["se_scode"] ?? "").Trim();
        se_branch = (Request["se_branch"] ?? "").Trim();//N
        se_branchnm = (Request["se_branchnm"] ?? "").Trim();
        se_dept = (Request["se_dept"] ?? "").Trim();
        se_syscode = (Request["se_syscode"] ?? "").Trim();
        se_seqtitle = (Request["se_seqtitle"] ?? "").Trim();

        qsort = (Request["qsort"] ?? "").Trim().ToLower();
        hend = (Request["hend"] ?? "").Trim();
        qcust_area = (Request["qcust_area"] ?? "").Trim();
        qcust_seq = (Request["qcust_seq"] ?? "").Trim();
        qscase_date = (Request["qscase_date"] ?? "").Trim();
        qecase_date = (Request["qecase_date"] ?? "").Trim();
        qseq = (Request["qseq"] ?? "").Trim();
        qseq1 = (Request["qseq1"] ?? "").Trim();
        qcust_prod = (Request["qcust_prod"] ?? "").Trim();
        qcustprod_no = (Request["qcustprod_no"] ?? "").Trim();
        qcase1 = (Request["qcase1"] ?? "").Trim();
        qscode1 = (Request["qscode1"] ?? "").Trim();
        
        try {
            WordOut();
        }
        finally {
            if (Rpt != null) Rpt.Dispose();
        }
    }

    protected void WordOut() {
        Dictionary<string, string> _tplFile = new Dictionary<string, string>();
        _tplFile.Add("csrpt", Server.MapPath("~/ReportTemplate/報表/國外專利案件進度報導.docx"));
        Rpt.CloneFromFile(_tplFile, true);

        string docFileName = string.Format("{0}國外專利案件進度報導.docx", se_scode);

        string SQL = "";
        string wSQL = "";
        if (qcust_area != "") wSQL += " and a.cust_area='" + qcust_area + "' ";
        if (qcust_seq != "") wSQL += " and a.cust_seq='" + qcust_seq + "' ";
        if (qscase_date != "") wSQL += " and a.in_date>='" + qscase_date + " 00:00:00' ";
        if (qecase_date != "") wSQL += " and a.in_date<='" + qecase_date + " 23:59:59' ";
        if (qseq != "") wSQL += " and a.seq='" + qseq + "' ";
        if (qseq1 != "") wSQL += " and a.seq1='" + qseq1 + "' ";
        if (qcust_prod != "") wSQL += " and a.cust_prod='" + qcust_prod + "' ";
        if (qcustprod_no != "") wSQL += " and a.custprod_no='" + qcustprod_no + "' ";
        if (qcase1 != "") wSQL += " and a.case1='" + qcase1 + "' ";
        if (hend == "Y") {//不含結案
            wSQL += " and (a.end_date is null or a.end_date='' or a.end_date>='" + DateTime.Now.ToString("yyyy/MM/dd") + "') ";
        }
        
        if (qscode1 != "") {
            wSQL += " and a.scode1='" + qscode1 + "' ";
        } else {
            //若為組主管，則求取可查詢之營洽名單
            if ((right & 128) > 0) {
            } else if ((right & 64) > 0) {
                string team_scode1 = "";
                SQL = "select distinct scode, sc_name,scode1 ";
                SQL += " from vscode_roles2 ";
                SQL += " WHERE dept='" + se_dept + "' AND Syscode='" + se_syscode + "' and branch='" + se_branch + "' and roles='sales' ";
		        SQL += " and master_scode='" +se_scode+ "'";
                SQL += " order by scode1,scode,sc_name ";
                using (DBHelper cnn = new DBHelper(Session["sysctrl"].ToString()).Debug(true)) {
                    SqlDataReader dr = cnn.ExecuteReader(SQL);
                    while (dr.Read()) {
                        team_scode1 += "," + dr.SafeRead("scode", "");
                    }
                }
                if (team_scode1 != "") team_scode1 = team_scode1.Substring(1);
                wSQL += " and a.scode1 in ('" + team_scode1.Replace(",", "','") + "') ";
            } else if ((right & 32) > 0) {
            }
        }
        
        using (DBHelper conn = new DBHelper(Session["btbrtdb"].ToString()).Debug(true)) {
            SQL="Select (Select sc_name from sysctrl.dbo.scode where scode=scode1) as Scodenm ";
            SQL = "Select a.*,b.rs_detail ";
            SQL += ",(select ap_cname1 from apcust where a.cust_seq=cust_seq and a.cust_area=cust_area) as ap_cname1 ";
            SQL += ",(select code_name from cust_code where code_type='PEcase_stat' and cust_code=a.now_stat) as case_statnm ";
            SQL += ",(select code_name from cust_code where code_type='PECase1' and cust_code=a.case1) as case1nm ";
            SQL += ",(select coun_c from sysctrl.dbo.country where coun_code=a.country) as countrynm ";
            SQL += ",(select sc_name from sysctrl.dbo.scode where scode=scode1) as Scodenm ";
            SQL += " from exp a,step_exp b";
            SQL += " where a.seq=b.seq and a.seq1=b.seq1 and a.step_grade=b.step_grade";
            SQL += wSQL;
            SQL += " order by " + qsort;

            DataTable dt = new DataTable();
            conn.DataTable(SQL, dt);
            Response.Write("件數="+dt.Rows.Count+"<BR>");
            
            //表頭
            Rpt.CopyBlock("b_title");
            Rpt.ReplaceBookmark("tdate", DateTime.Now.ToShortDateString());
            
            if (hend == "Y") {
                Rpt.ReplaceBookmark("end", "(不含結案案件)");
            } else {
                Rpt.ReplaceBookmark("end", "(含已結案案件)");
            }

            if (qsort == "a.seq,a.seq1") {
                Rpt.ReplaceBookmark("sort", "依本所編號");
            } else if (qsort == "a.cust_prod") {
                Rpt.ReplaceBookmark("sort", "依客戶卷號");
            } else if (qsort == "a.custprod_no") {
                Rpt.ReplaceBookmark("sort", "依客戶商品號");
            }

            Rpt.ReplaceBookmark("custcount", dt.Rows.Count.ToString());
            
            if (dt.Rows.Count>0){
                Rpt.ReplaceBookmark("ap_cname", dt.Rows[0]["ap_cname1"].ToString());
            }
            string branchtitle = se_branchnm.Substring(0, 2) + "營業所專利部";
            if (dt.Rows.Count>0){
                SQL = "select pscode,(select sc_name from sysctrl.dbo.scode where scode=a.pscode) as sc_name ";
                SQL += " from custz a where cust_area='" + dt.Rows[0]["cust_area"] + "' and cust_seq=" + dt.Rows[0]["cust_seq"];
                using (SqlDataReader dr= conn.ExecuteReader(SQL)){
                    if (dr.Read()) {
                        branchtitle += dr.SafeRead("sc_name", "");
                    }
                }
            }
            Rpt.ReplaceBookmark("scode", branchtitle);

            for (int i = 0; i < dt.Rows.Count; i++) {
                int pno = i + 1;
                Rpt.CopyBlock("b_detail");
                //聖島編號(國別)
                string seq = se_branch + "PE" + dt.Rows[i].SafeRead("seq", "");
                if (dt.Rows[i].SafeRead("seq1", "") != "_")
                    seq += "_" + dt.Rows[i].SafeRead("seq1", "");
                if (seq != "")
                    seq += "(" + dt.Rows[i].SafeRead("countrynm", "") + ")";
                Rpt.ReplaceText("#seq#", seq);
                //案件名稱
                Rpt.ReplaceBookmark("cappl_name", dt.Rows[i].SafeRead("cappl_name", "").ToXmlUnicode());
                //國外代理人聯繫案件編號(國外所編號)
                string ftp_no = "";
                if (dt.Rows[i].SafeRead("tp_no", "") != "" && dt.Rows[i].SafeRead("tp_no", "") != "0") {
                    ftp_no = se_seqtitle + "-" + dt.Rows[i].SafeRead("tp_no", "");
                    if (dt.Rows[i].SafeRead("tp_no1", "") != "_") {
                        ftp_no += "-" + dt.Rows[i].SafeRead("tp_no1", "");
                    }
                    if (dt.Rows[i].SafeRead("country", "") != "") {
                        ftp_no += " " + dt.Rows[i].SafeRead("country", "").ToUpper();
                    }
                }
                Rpt.ReplaceBookmark("ftp_no", ftp_no);
                //專利案性
                Rpt.ReplaceBookmark("case1nm", dt.Rows[i].SafeRead("case1nm", ""));
                //案件狀態
                Rpt.ReplaceBookmark("case_statnm", dt.Rows[i].SafeRead("case_statnm", ""));
                //應辦期限minctrl_date
                SQL = "select seq,seq1,ctrl_type,MIN(ctrl_date) minctrl_date ";
                SQL += "from ctrl_exp where ctrl_type in('A1','B1') ";//法定期限&自管期限
                SQL += "and seq='" + dt.Rows[i]["seq"] + "' and seq1='" + dt.Rows[i]["seq1"] + "' ";
                SQL += "group by seq,seq1,ctrl_type ";
                SQL += "order by ctrl_type ";
                string minctrl_date = "";
                using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
                    if (dr.Read()) {
                        minctrl_date = dr.GetDateTimeString("minctrl_date", "yyyy/M/d");
                    }
                }
                Rpt.ReplaceBookmark("minctrl_date", minctrl_date);
                //貴方卷號
                Rpt.ReplaceBookmark("cust_prod", dt.Rows[i].SafeRead("cust_prod", ""));
                //客戶產品編號
                Rpt.ReplaceBookmark("custprod_no", dt.Rows[i].SafeRead("custprod_no", ""));
                //申請人ap_cname1
                string ap_cname = "";
                SQL = "Select * from ap_exp where seq='" + dt.Rows[i]["seq"] + "' and seq1='" + dt.Rows[i]["seq1"] + "'";
                using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
                    while (dr.Read()) {
                        ap_cname += "、" + dr.SafeRead("ap_cname1", "").Trim() + dr.SafeRead("ap_cname2", "").Trim();
                    }
                }
                if (ap_cname != "")
                    ap_cname = ap_cname.Substring(1);
                Rpt.ReplaceBookmark("ap_cname1", ap_cname);
                //發明/創作人
                string ant_cname = "";
                SQL = "Select * from ap_expant where seq='" + dt.Rows[i]["seq"] + "' and seq1='" + dt.Rows[i]["seq1"] + "'";
                using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
                    while (dr.Read()) {
                        ant_cname += "、" + dr.SafeRead("ant_cname1", "").Trim() + dr.SafeRead("ant_cname2", "").Trim();
                    }
                }
                if (ant_cname != "")
                    ant_cname = ant_cname.Substring(1);
                Rpt.ReplaceBookmark("ant_cname", ant_cname);
                //申請日
                Rpt.ReplaceBookmark("apply_date", dt.Rows[i].GetDateTimeString("apply_date", "yyyy/M/d"));
                //申請案號
                Rpt.ReplaceBookmark("apply_no", dt.Rows[i].SafeRead("apply_no", ""));
                //發證日期
                Rpt.ReplaceBookmark("issue_date", dt.Rows[i].GetDateTimeString("issue_date", "yyyy/M/d"));
                //早期公開日
                Rpt.ReplaceBookmark("pub_date", dt.Rows[i].GetDateTimeString("pub_date", "yyyy/M/d"));
                //早期公開號
                Rpt.ReplaceBookmark("pub_no", dt.Rows[i].SafeRead("pub_no", ""));
                //證書號碼
                Rpt.ReplaceBookmark("issue_no", dt.Rows[i].SafeRead("issue_no", ""));
                //公告日期
                Rpt.ReplaceBookmark("open_date", dt.Rows[i].GetDateTimeString("open_date", "yyyy/M/d"));
                //公告號碼
                Rpt.ReplaceBookmark("open_no", dt.Rows[i].SafeRead("open_no", ""));
                //專利權期間
                string term = dt.Rows[i].GetDateTimeString("term1", "yyyy/M/d") + "~" + dt.Rows[i].GetDateTimeString("term2", "yyyy/M/d");
                Rpt.ReplaceBookmark("term", term);
                //已繳年數
                Rpt.ReplaceBookmark("pay_times", dt.Rows[i].SafeRead("pay_times", ""));
                //續繳年費期限
                string minpay_date = "";
                SQL = "SELECT MIN(pay_date) AS minpay_date FROM ann_exp where seq='" + dt.Rows[i]["seq"] + "' and seq1='" + dt.Rows[i]["seq1"] + "'";
                using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
                    if (dr.Read()) {
                        minpay_date = dr.GetDateTimeString("minpay_date", "yyyy/M/d");
                    }
                }
                Rpt.ReplaceBookmark("minpay_date", minpay_date);
                //結案日期
                Rpt.ReplaceBookmark("End_date", dt.Rows[i].GetDateTimeString("End_date", "yyyy/M/d"));

                //第一頁是3筆換頁,其他是每4筆換頁
                if (pno == 3) {
                    Rpt.NewPage();
                    Rpt.AddParagraph();
                } else if ((pno > 3 && (pno - 3) % 4 == 0) && pno != dt.Rows.Count) {
                    Rpt.NewPage();
                    Rpt.AddParagraph();
                }
            }
        }
        Rpt.CopyPageFoot("csrpt", false);//複製頁尾/邊界
        Rpt.Flush(docFileName);
    }
</script>
