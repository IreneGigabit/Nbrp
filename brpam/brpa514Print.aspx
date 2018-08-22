<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>

<script runat="server">
    protected OpenXmlHelper Rpt = new OpenXmlHelper();
    protected string dept="";
    protected string branch = "";
    protected string send_way = "";

    private void Page_Load(System.Object sender, System.EventArgs e) {
        Response.CacheControl = "Private";
        Response.AddHeader("Pragma", "no-cache");
        Response.Expires = -1;
        Response.Clear();

        dept = (Request["dept"] ?? "").ToString();//P
        branch = (Request["cust_area"] ?? "").ToString();//N
        send_way = (Request["send_way"] ?? "").ToString();//E

        try {
            WordOut();
        }
        finally {
            if (Rpt != null) Rpt.Dispose();
        }
    }

    protected void WordOut() {
        Dictionary<string, string> _tplFile = new Dictionary<string, string>();
        _tplFile.Add("gsrpt", Server.MapPath("~/ReportTemplate/報表/發文回條.docx"));
        Rpt.CloneFromFile(_tplFile, true);

        string docFileName = string.Format("GS{0}-514P-{1:yyyyMMdd}.docx", send_way, DateTime.Today);

        string SQL = "";
        using (DBHelper conn = new DBHelper(Session["btbrtdb"].ToString()).Debug(true)) {
            SQL = "select dmp_scode,cappl_name,apply_no,change_no,issue_no,open_no,term1,term2,new,case_no,";
            SQL += "rs_no,branch,seq,seq1,step_grade,step_date,send_way,mp_date,send_selnm,send_clnm,send_cl1nm,rs_class,rs_detail,fees,pr_scode,";
            SQL += "a.send_sel,(select mark1 from cust_code where code_type='SEND_SEL' and cust_code=a.send_sel) as send_selfel,"; //發文性質的欄位名稱
            SQL += "(select branchname from sysctrl.dbo.branch_code where branch=a.branch) as branchnm,mapply_no,";
            SQL += "receipt_type,receipt_title,rectitle_name ";
            SQL += " from vstep_dmp a where cg='G' and rs='S' ";
            SQL += " and a.cancel_flag<>'Y' ";
            //20170605 因應電子收據上線，不顯示電子收據資料
            SQL += " and isnull(receipt_type,'')<>'E' ";
            if ((Request["send_way"] ?? "") != "") SQL += " and send_way='" + Request["send_way"] + "'";
            if ((Request["sdate"] ?? "") != "") SQL += " and step_date>='" + Request["sdate"] + "'";
            if ((Request["edate"] ?? "") != "") SQL += " and step_date<='" + Request["edate"] + "'";
            if ((Request["srs_no"] ?? "") != "") SQL += " and rs_no>='" + Request["srs_no"] + "'";
            if ((Request["ers_no"] ?? "") != "") SQL += " and rs_no<='" + Request["ers_no"] + "'";
            if ((Request["seq"] ?? "") != "") SQL += " and seq=" + Request["seq"];
            if ((Request["sseq"] ?? "") != "") SQL += " and seq>=" + Request["sseq"];
            if ((Request["eseq"] ?? "") != "") SQL += " and seq<=" + Request["eseq"];
            if ((Request["seq1"] ?? "") != "") SQL += " and seq1='" + Request["seq1"] + "'";
            if ((Request["dmp_scode"] ?? "") != "") SQL += " and dmp_scode='" + Request["dmp_scode"] + "'";
            if ((Request["scust_seq"] ?? "") != "") SQL += " and cust_seq>=" + Request["scust_seq"];
            if ((Request["ecust_seq"] ?? "") != "") SQL += " and cust_seq<=" + Request["ecust_seq"];
            if ((Request["rs_type"] ?? "") != "") SQL += " and rs_type='" + Request["rs_type"] + "'";
            if ((Request["rs_class"] ?? "") != "") SQL += " and rs_class='" + Request["rs_class"] + "'";
            if ((Request["rs_code"] ?? "") != "") SQL += " and rs_code='" + Request["rs_code"] + "'";
            if ((Request["act_code"] ?? "") != "") SQL += " and act_code='" + Request["act_code"] + "'";
            if ((Request["step_grade"] ?? "") != "") SQL += " and step_grade='" + Request["step_grade"] + "'";
            if ((Request["hprint"] ?? "") == "N") {
                SQL += " and (substring(new,2,1)='" + Request["hprint"] + "'";
                SQL += " or substring(new,2,1)='')";
            }
            SQL += " order by a.rs_no";

            DataTable dt = new DataTable();
            conn.DataTable(SQL, dt);

            for (int i = 0; i < dt.Rows.Count; i++) {
                int runTime = 1;
                //副本有選擇要多一印份給副本收文者
                if (dt.Rows[i].SafeRead("send_cl1nm", "") != "") runTime = 2;

                for (int r = 1; r <= runTime; r++) {
                    Rpt.CopyBlock("b_table");
                    //總管處發文日期
                    DateTime mp_date;
                    string mpDate = DateTime.TryParse(dt.Rows[i].SafeRead("mp_date", ""), out mp_date) ? mp_date.ToShortDateString() : "";
                    Rpt.ReplaceBookmark("mp_date", mpDate);
                    
                    //發文序號
                    string strrs_no = string.Format("發文({0})聖{1}{2}　{3}　字第　{4}　號"
                                        , DateTime.TryParse(mpDate, out mp_date) ? (mp_date.Year - 1911).ToString() : ""
                                        , dt.Rows[i].SafeRead("branchnm", "").Substring(1, 1)
                                        , (dept.ToUpper() == "T" ? "商" : "") + (dept.ToUpper() == "P" ? "專" : "")
                                        , dt.Rows[i].SafeRead("pr_scode", "")
                                        , dt.Rows[i].SafeRead("rs_no", "").Substring(2)
                                        );
                    Rpt.ReplaceBookmark("strrs_no", strrs_no);
                    
                    //受文者，發文單位
                    if (r == 2) {
                        Rpt.ReplaceBookmark("send_clnm", "副本\n" + dt.Rows[i].SafeRead("send_cl1nm", ""));
                    } else {
                        Rpt.ReplaceBookmark("send_clnm", dt.Rows[i].SafeRead("send_clnm", ""));
                    }
                    
                    //簡由，發文性質+案件名稱+發文內容
                    string send_detail = "";
                    string send_sel = dt.Rows[i].SafeRead("send_sel", "").Trim();
                    string str1 = "";
                    if (send_sel != "") {
                        switch (send_sel) {
                            case "1":
                                if (dt.Rows[i].SafeRead("apply_no", "").Trim() != "")
                                    str1 = "申請號 第　" + dt.Rows[i].SafeRead("apply_no", "").Trim() + "　號";
                                break;
                            case "2":
                                if (dt.Rows[i].SafeRead("change_no", "").Trim() != "")
                                    str1 = "改請號 第　" + dt.Rows[i].SafeRead("change_no", "").Trim() + "　號";
                                break;
                            case "3":
                                if (dt.Rows[i].SafeRead("issue_no", "").Trim() != "")
                                    str1 = "公告號 第　" + dt.Rows[i].SafeRead("issue_no", "").Trim() + "　號";
                                break;
                            case "4":
                                if (dt.Rows[i].SafeRead("open_no", "").Trim() != "")
                                    str1 = "公開號 第　" + dt.Rows[i].SafeRead("open_no", "").Trim() + "　號";
                                break;
                        }
                    } else {
                        if (dt.Rows[i].SafeRead("change_no", "").Trim() != "") {
                            str1 = "改請號 第　" + dt.Rows[i].SafeRead("change_no", "").Trim() + "　號";
                        } else {
                            if (dt.Rows[i].SafeRead("apply_no", "").Trim() != "")
                                str1 = "申請號 第　" + dt.Rows[i].SafeRead("apply_no", "").Trim() + "　號";
                        }
                    }
                    if (send_detail != "" && str1 != "") send_detail += "\n";
                    send_detail += str1;
                        
                    string str2 = dt.Rows[i].SafeRead("cappl_name", "").Trim();
                    if (send_detail != "" && str2 != "") send_detail += "\n";
                    send_detail += str2;
                    
                    string str3 = dt.Rows[i].SafeRead("rs_detail", "").Trim();
                    if (send_detail != "" && str3 != "") send_detail += "\n";
                    send_detail += str3;
                    
                    string str4 = "";
                    if (dt.Rows[i].SafeRead("mapply_no", "").Trim() != "") {
                        str4 = "(母案申請號 第　" + dt.Rows[i].SafeRead("mapply_no", "").Trim() + "　號)";
                    }
                    if (send_detail != "" && str4 != "") send_detail += "\n";
                    send_detail += str4;

                    string str5 = "";
                    if (dt.Rows[i].SafeRead("rs_class", "").Trim() == "P50") {
                        str5 = "專用期限:" + dt.Rows[i].SafeRead("term1", "") + " ~ " + dt.Rows[i].SafeRead("term2", "");
                    }
                    if (send_detail != "" && str5 != "") send_detail += "\n";
                    send_detail += str5;
                    
                    string str6 = "";
                    if (dt.Rows[i].SafeRead("send_way", "").Trim() == "E") {
                        str6 = "※電子年費送件";
                    } else if (dt.Rows[i].SafeRead("send_way", "").Trim() == "ES") {
                        str6 = "※一般電子送件";
                    }
                    if (send_detail != "" && str6 != "") send_detail += "\n";
                    send_detail += str6;
                    
                    //20180621 增加收據抬頭,若是空白則不顯示
			        //有指定抬頭才要顯示
                    string receipt_title=dt.Rows[i].SafeRead("receipt_title", "");
                    if (receipt_title == "A" || receipt_title == "C") {
                        string rectitle_name = dt.Rows[i].SafeRead("rectitle_name", "");
                        //若step_dmp裡面沒有記到的話回頭去抓交辦資料
                        if (rectitle_name == "") {
                            SQL = "select distinct b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2 ";
                            SQL += " ,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename ";
                            SQL += " from dmp_apcust a ";
                            SQL += " inner join apcust b on a.apsqlno=b.apsqlno ";
                            SQL += " where a.kind in('A') ";
                            SQL += " and exists (select 1 from case_dmp c where case_no in('" + dt.Rows[i].SafeRead("case_no", "").Replace(",", "','") + "') and case_no<>'' ";
                            SQL += " 			and a.in_no=c.in_no and a.in_scode=c.in_scode) ";
                            using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
                                while (dr.Read()) {
                                    string Cname_string = "";
                                    //本國公司
                                    if (dr.SafeRead("apclass", "").Left(1) == "A") {
                                        Cname_string = dr.SafeRead("ap_cname1", "").Trim() + dr.SafeRead("ap_cname2", "").Trim();
                                    }

                                    //本國自然人
                                    if (dr.SafeRead("apclass", "").Left(1) == "B") {
                                        Cname_string = dr.SafeRead("ap_fcname", "").Trim() + dr.SafeRead("ap_lcname", "").Trim();
                                        if (Cname_string == "")
                                            Cname_string = dr.SafeRead("ap_cname1", "").Trim() + dr.SafeRead("ap_cname2", "").Trim();
                                    }

                                    //20161206外國人/公司增加判斷-若有填寫申請人姓&名,則顯示姓名
                                    if (dr.SafeRead("apclass", "").Left(1) == "C") {
                                        Cname_string = dr.SafeRead("ap_fcname", "").Trim() + dr.SafeRead("ap_lcname", "").Trim();
                                        if (Cname_string == "")
                                            Cname_string = dr.SafeRead("ap_cname1", "").Trim() + dr.SafeRead("ap_cname2", "").Trim();
                                    }

                                    if (rectitle_name != "") rectitle_name += "、";
                                    rectitle_name += Cname_string;
                                }
                            }
                        }

                        if (send_detail != "") send_detail += "\n";
                        if (receipt_title == "C")//專利權人(代繳人)
                            rectitle_name = rectitle_name + "(代繳人：聖島國際專利商標聯合事務所)";
                        send_detail += "收據抬頭：" + rectitle_name;
                    }
                    Rpt.ReplaceBookmark("send_detail", send_detail);
                    
                    //本所編號
                    string seq = branch + dept + dt.Rows[i].SafeRead("seq", "");
                    if (dt.Rows[i].SafeRead("seq1", "_") != "_")
                        seq = dt.Rows[i].SafeRead("seq", "") + "-" + dt.Rows[i].SafeRead("seq1", "");
                    Rpt.ReplaceBookmark("seq", seq);
                    
                    //最後期限，法定期限(本次官發銷管的管制日期)
                    string ctrl_date = "";
                    SQL = "select ctrl_date from resp_dmp where branch='" + dt.Rows[i]["branch"] + "' and seq=" + dt.Rows[i]["seq"];
                    SQL += " and seq1='" + dt.Rows[i]["seq1"] + "' and resp_grade=" + dt.Rows[i]["step_grade"] + " and substring(ctrl_type,1,1)='A'";
                    using (SqlDataReader dr = conn.ExecuteReader(SQL)) {
                        while (dr.Read()) {
                            if (!(dr["ctrl_date"] is DBNull) && dr["ctrl_date"] != "") {
                                if (ctrl_date != "") ctrl_date += "\n";
                                ctrl_date += Convert.ToDateTime(dr["ctrl_date"]).ToShortDateString();
                            }
                        }
                    }
                    Rpt.ReplaceBookmark("ctrl_date", ctrl_date);
                    
                    //規費
                    if (r == 2)
                        Rpt.ReplaceBookmark("fees", "0 元");
                    else
                        Rpt.ReplaceBookmark("fees", dt.Rows[i]["fees"] + " 元");
                }
            }
        }
        Rpt.CopyPageFoot("gsrpt", false);//複製頁尾/邊界
        Rpt.Flush(docFileName);
        //Rpt.SaveTo(Server.MapPath("~/reportdata/" + docFileName));
        //ipoRpt.SetPrint();
    }

		
	
</script>
