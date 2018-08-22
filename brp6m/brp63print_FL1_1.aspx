<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>

<script runat="server">
	protected IPOReport ipoRpt = null;

	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;
		Response.Clear();

        string in_scode = (Request["in_scode"] ?? "").ToString();//n100
        string in_no = (Request["in_no"] ?? "").ToString();//20170103001
        string branch = (Request["branch"] ?? "").ToString();//N
		string rectitle = (Request["receipt_title"] ?? "").ToString();//N

		try {
            ipoRpt = new IPOReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch, rectitle)
            {
                ReportCode = "FL1_1",
            }.Init();
			WordOut();
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		Dictionary<string, string> _tplFile = new Dictionary<string, string>();
        _tplFile.Add("apply", Server.MapPath("~/ReportTemplate/申請書/51[專簡B]專利權授權登記申請書FL1_1.docx"));
		_tplFile.Add("base", Server.MapPath("~/ReportTemplate/申請書/00基本資料表.docx"));
		ipoRpt.CloneFromFile(_tplFile, true);

        string docFileName = ipoRpt.Seq + "-專利權授權.docx";
		
		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			//標題區塊
			ipoRpt.CopyBlock("b_title");
            //專利類別
            if (dmp.Rows[0]["s_case1nm"].ToString() != "") {
                ipoRpt.ReplaceBookmark("case1nm", dmp.Rows[0]["s_case1nm"].ToString());
            } else {
                ipoRpt.ReplaceBookmark("case1nm", "發明／新型／設計");
            }
            //申請案號
			if (dmp.Rows[0]["change_no"].ToString() != "") {
				ipoRpt.ReplaceBookmark("apply_no", dmp.Rows[0]["change_no"].ToString());
			} else {
				ipoRpt.ReplaceBookmark("apply_no", dmp.Rows[0]["apply_no"].ToString());
			}
            //專利證書號數
            ipoRpt.ReplaceBookmark("capply_no", dmp.Rows[0]["capply_no"].ToString());
            //事務所或申請人案件編號
			ipoRpt.ReplaceBookmark("seq", ipoRpt.Seq + "-" + dmp.Rows[0]["scode1"].ToString());
			//授權人
			using (DataTable dtAp = ipoRpt.GetApAnt("D1")) {
                if (dtAp.Rows.Count > 0) {
                    for (int i = 0; i < dtAp.Rows.Count; i++) {
                        ipoRpt.CopyBlock("b_apply");
                        ipoRpt.ReplaceBookmark("apply_type", "授權人");
                        ipoRpt.ReplaceBookmark("apply_num", (i + 1).ToString());
                        ipoRpt.ReplaceBookmark("ap_country", dtAp.Rows[i]["Country_name"].ToString());
                        ipoRpt.ReplaceBookmark("ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
                        ipoRpt.ReplaceBookmark("ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
                        ipoRpt.ReplaceBookmark("ap_cname", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
                        ipoRpt.ReplaceBookmark("ap_ename", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
                    }
                } else {
                    ipoRpt.CopyBlock("b_apply");
                    ipoRpt.ReplaceBookmark("apply_type", "授權人");
                    ipoRpt.ReplaceBookmark("apply_num", "1");
                    ipoRpt.ReplaceBookmark("ap_country", "");
                    ipoRpt.ReplaceBookmark("ap_cname_title", "中文名稱／中文姓名");
                    ipoRpt.ReplaceBookmark("ap_ename_title", "英文名稱／英文姓名");
                    ipoRpt.ReplaceBookmark("ap_cname", "");
                    ipoRpt.ReplaceBookmark("ap_ename", "");
                }
			}
            //授權人之代理人
			ipoRpt.CopyBlock("b_agent");
			using (DataTable dtAgt = ipoRpt.Agent) {
                ipoRpt.ReplaceBookmark("agt_type1", "授權人");
                ipoRpt.ReplaceBookmark("agt_name1", dtAgt.Rows[0]["agt_name1"].ToString().Trim());
                ipoRpt.ReplaceBookmark("agt_type2", "授權人");
                ipoRpt.ReplaceBookmark("agt_name2", dtAgt.Rows[0]["agt_name2"].ToString().Trim());
			}
            //被授權人
            using (DataTable dtAp = ipoRpt.GetApAnt("D2")) {
                for (int i = 0; i < dtAp.Rows.Count; i++) {
                    ipoRpt.CopyBlock("b_apply");
                    ipoRpt.ReplaceBookmark("apply_type", "被授權人");
                    ipoRpt.ReplaceBookmark("apply_num", (i + 1).ToString());
                    ipoRpt.ReplaceBookmark("ap_country", dtAp.Rows[i]["Country_name"].ToString());
                    ipoRpt.ReplaceBookmark("ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
                    ipoRpt.ReplaceBookmark("ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
                    ipoRpt.ReplaceBookmark("ap_cname", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
                    ipoRpt.ReplaceBookmark("ap_ename", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
                }
            }
            //被授權人之代理人
            ipoRpt.CopyBlock("b_agent");
            using (DataTable dtAgt = ipoRpt.Agent) {
                ipoRpt.ReplaceBookmark("agt_type1", "被授權人");
                ipoRpt.ReplaceBookmark("agt_name1", dtAgt.Rows[0]["agt_name1"].ToString().Trim());
                ipoRpt.ReplaceBookmark("agt_type2", "被授權人");
                ipoRpt.ReplaceBookmark("agt_name2", dtAgt.Rows[0]["agt_name2"].ToString().Trim());
            }
            //同時辦理事項/繳費資訊
			ipoRpt.CopyBlock("b_content");
			ipoRpt.ReplaceBookmark("receipt_name", ipoRpt.RectitleName);
			//附送書件
			ipoRpt.CopyReplaceBlock("b_attach", "#seq#", ipoRpt.Seq);
			//ipoRpt.CopyReplaceBlock("b_attach", new Dictionary<string, string>() { { "#seq#", ipoRpt.Seq }, { "#case1nm#", Case1nm } });
			//具結
			ipoRpt.CopyBlock("b_sign");

			bool baseflag = true;//是否產生基本資料表
			ipoRpt.CopyPageFoot("apply", baseflag);//申請書頁尾
			if (baseflag) {
                ipoRpt.CopyBlock("base", "base_title");
                //授權人
                using (DataTable Apcust = ipoRpt.GetApAnt("D1")) {
                    if (Apcust.Rows.Count > 0) {
                        for (int i = 0; i < Apcust.Rows.Count; i++) {
                            ipoRpt.CopyBlock("base", "base_apcust");
                            ipoRpt.ReplaceBookmark("base_ap_type", "授權人");
                            ipoRpt.ReplaceBookmark("base_ap_num", (i + 1).ToString());
                            ipoRpt.ReplaceBookmark("base_ap_country", Apcust.Rows[i]["Country_name"].ToString());
                            ipoRpt.ReplaceBookmark("ap_class", Apcust.Rows[i]["apclass_name"].ToString());
                            if (Apcust.Rows[i]["ap_country"].ToString() == "T") {
                                ipoRpt.ReplaceBookmark("apcust_no", Apcust.Rows[i]["apcust_no"].ToString());
                            } else {
                                ipoRpt.ReplaceBookmark("apcust_no", "", true);
                            }
                            ipoRpt.ReplaceBookmark("base_ap_cname_title", Apcust.Rows[i]["Title_cname"].ToString());
                            ipoRpt.ReplaceBookmark("base_ap_ename_title", Apcust.Rows[i]["Title_ename"].ToString());
                            ipoRpt.ReplaceBookmark("base_ap_cname", Apcust.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
                            ipoRpt.ReplaceBookmark("base_ap_ename", Apcust.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
                            ipoRpt.ReplaceBookmark("ap_live_country", Apcust.Rows[i]["Country_name"].ToString());
                            ipoRpt.ReplaceBookmark("ap_zip", Apcust.Rows[i]["ap_zip"].ToString());
                            string ap_addr = Apcust.Rows[i]["ap_addr1"].ToString().ToXmlUnicode()
                                + Apcust.Rows[i]["ap_addr2"].ToString().ToXmlUnicode();
                            ipoRpt.ReplaceBookmark("ap_addr", ap_addr);
                            string ap_eddr = Apcust.Rows[i]["ap_eaddr1"].ToString().ToXmlUnicode(true)
                                + Apcust.Rows[i]["ap_eaddr2"].ToString().ToXmlUnicode(true)
                                + Apcust.Rows[i]["ap_eaddr3"].ToString().ToXmlUnicode(true)
                                + Apcust.Rows[i]["ap_eaddr4"].ToString().ToXmlUnicode(true);
                            ipoRpt.ReplaceBookmark("ap_eddr", ap_eddr);
                            ipoRpt.ReplaceBookmark("ap_crep", Apcust.Rows[i]["ap_crep"].ToString().ToXmlUnicode());
                            ipoRpt.ReplaceBookmark("ap_erep", Apcust.Rows[i]["ap_erep"].ToString().ToXmlUnicode(true));
                        }
                    } else {
                        ipoRpt.CopyBlock("base", "base_apcust");
                        ipoRpt.ReplaceBookmark("base_ap_type", "授權人");
                        ipoRpt.ReplaceBookmark("base_ap_num", "1");
                        ipoRpt.ReplaceBookmark("base_ap_country", "");
                        ipoRpt.ReplaceBookmark("ap_class", "");
                        ipoRpt.ReplaceBookmark("apcust_no", "");
                        ipoRpt.ReplaceBookmark("base_ap_cname_title", "中文名稱／中文姓名");
                        ipoRpt.ReplaceBookmark("base_ap_ename_title", "英文名稱／英文姓名");
                        ipoRpt.ReplaceBookmark("base_ap_cname", "");
                        ipoRpt.ReplaceBookmark("base_ap_ename", "");
                        ipoRpt.ReplaceBookmark("ap_live_country", "");
                        ipoRpt.ReplaceBookmark("ap_zip", "");
                        ipoRpt.ReplaceBookmark("ap_addr", "");
                        ipoRpt.ReplaceBookmark("ap_eddr", "");
                        ipoRpt.ReplaceBookmark("ap_crep", "");
                        ipoRpt.ReplaceBookmark("ap_erep", "");
                    }
                }
                //授權人之代理人
                using (DataTable Agent = ipoRpt.Agent) {
                    for (int i = 0; i < Agent.Rows.Count; i++) {
                        ipoRpt.CopyBlock("base", "base_agent");
                        ipoRpt.ReplaceBookmark("agt_type1", "授權人之");
                        ipoRpt.ReplaceBookmark("agt_idno1", Agent.Rows[i]["agt_idno1"].ToString());
                        ipoRpt.ReplaceBookmark("agt_id1", Agent.Rows[i]["agt_id1"].ToString());
                        ipoRpt.ReplaceBookmark("base_agt_name1", Agent.Rows[i]["agt_name1"].ToString());
                        ipoRpt.ReplaceBookmark("agt_zip1", Agent.Rows[i]["agt_zip"].ToString());
                        ipoRpt.ReplaceBookmark("agt_addr1", Agent.Rows[i]["agt_addr"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_tel1", Agent.Rows[i]["agt_tel"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_fax1", Agent.Rows[i]["agt_fax"].ToString());

                        ipoRpt.ReplaceBookmark("agt_type2", "授權人之");
                        ipoRpt.ReplaceBookmark("agt_idno2", Agent.Rows[i]["agt_idno2"].ToString());
                        ipoRpt.ReplaceBookmark("agt_id2", Agent.Rows[i]["agt_id2"].ToString());
                        ipoRpt.ReplaceBookmark("base_agt_name2", Agent.Rows[i]["agt_name2"].ToString());
                        ipoRpt.ReplaceBookmark("agt_zip2", Agent.Rows[i]["agt_zip"].ToString());
                        ipoRpt.ReplaceBookmark("agt_addr2", Agent.Rows[i]["agt_addr"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_tel2", Agent.Rows[i]["agt_tel"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_fax2", Agent.Rows[i]["agt_fax"].ToString());
                    }
                }

                //被授權人
                using (DataTable Apcust = ipoRpt.GetApAnt("D2")) {
                    for (int i = 0; i < Apcust.Rows.Count; i++) {
                        ipoRpt.CopyBlock("base", "base_apcust");
                        ipoRpt.ReplaceBookmark("base_ap_type", "被授權人");
                        ipoRpt.ReplaceBookmark("base_ap_num", (i + 1).ToString());
                        ipoRpt.ReplaceBookmark("base_ap_country", Apcust.Rows[i]["Country_name"].ToString());
                        ipoRpt.ReplaceBookmark("ap_class", Apcust.Rows[i]["apclass_name"].ToString());
                        if (Apcust.Rows[i]["ap_country"].ToString() == "T") {
                            ipoRpt.ReplaceBookmark("apcust_no", Apcust.Rows[i]["apcust_no"].ToString());
                        } else {
                            ipoRpt.ReplaceBookmark("apcust_no", "", true);
                        }
                        ipoRpt.ReplaceBookmark("base_ap_cname_title", Apcust.Rows[i]["Title_cname"].ToString());
                        ipoRpt.ReplaceBookmark("base_ap_ename_title", Apcust.Rows[i]["Title_ename"].ToString());
                        ipoRpt.ReplaceBookmark("base_ap_cname", Apcust.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
                        ipoRpt.ReplaceBookmark("base_ap_ename", Apcust.Rows[i]["Ename_string"].ToString().ToXmlUnicode(true));
                        ipoRpt.ReplaceBookmark("ap_live_country", Apcust.Rows[i]["Country_name"].ToString());
                        ipoRpt.ReplaceBookmark("ap_zip", Apcust.Rows[i]["ap_zip"].ToString());
                        string ap_addr = Apcust.Rows[i]["ap_addr1"].ToString().ToXmlUnicode()
                            + Apcust.Rows[i]["ap_addr2"].ToString().ToXmlUnicode();
                        ipoRpt.ReplaceBookmark("ap_addr", ap_addr);
                        string ap_eddr = Apcust.Rows[i]["ap_eaddr1"].ToString().ToXmlUnicode(true)
                            + Apcust.Rows[i]["ap_eaddr2"].ToString().ToXmlUnicode(true)
                            + Apcust.Rows[i]["ap_eaddr3"].ToString().ToXmlUnicode(true)
                            + Apcust.Rows[i]["ap_eaddr4"].ToString().ToXmlUnicode(true);
                        ipoRpt.ReplaceBookmark("ap_eddr", ap_eddr);
                        ipoRpt.ReplaceBookmark("ap_crep", Apcust.Rows[i]["ap_crep"].ToString().ToXmlUnicode());
                        ipoRpt.ReplaceBookmark("ap_erep", Apcust.Rows[i]["ap_erep"].ToString().ToXmlUnicode(true));
                    }
                }
                //被授權人之代理人
                using (DataTable Agent = ipoRpt.Agent) {
                    for (int i = 0; i < Agent.Rows.Count; i++) {
                        ipoRpt.CopyBlock("base", "base_agent");
                        ipoRpt.ReplaceBookmark("agt_type1", "被授權人之");
                        ipoRpt.ReplaceBookmark("agt_idno1", Agent.Rows[i]["agt_idno1"].ToString());
                        ipoRpt.ReplaceBookmark("agt_id1", Agent.Rows[i]["agt_id1"].ToString());
                        ipoRpt.ReplaceBookmark("base_agt_name1", Agent.Rows[i]["agt_name1"].ToString());
                        ipoRpt.ReplaceBookmark("agt_zip1", Agent.Rows[i]["agt_zip"].ToString());
                        ipoRpt.ReplaceBookmark("agt_addr1", Agent.Rows[i]["agt_addr"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_tel1", Agent.Rows[i]["agt_tel"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_fax1", Agent.Rows[i]["agt_fax"].ToString());

                        ipoRpt.ReplaceBookmark("agt_type2", "被授權人之");
                        ipoRpt.ReplaceBookmark("agt_idno2", Agent.Rows[i]["agt_idno2"].ToString());
                        ipoRpt.ReplaceBookmark("agt_id2", Agent.Rows[i]["agt_id2"].ToString());
                        ipoRpt.ReplaceBookmark("base_agt_name2", Agent.Rows[i]["agt_name2"].ToString());
                        ipoRpt.ReplaceBookmark("agt_zip2", Agent.Rows[i]["agt_zip"].ToString());
                        ipoRpt.ReplaceBookmark("agt_addr2", Agent.Rows[i]["agt_addr"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_tel2", Agent.Rows[i]["agt_tel"].ToString());
                        ipoRpt.ReplaceBookmark("agatt_fax2", Agent.Rows[i]["agt_fax"].ToString());
                    }
                }

                ipoRpt.CopyPageFoot("base", false);//頁尾
			}
		}

		ipoRpt.Flush(docFileName);
		ipoRpt.SetPrint();
	}
</script>
