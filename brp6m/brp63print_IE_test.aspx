<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>
<%@ Import Namespace = "DocumentFormat.OpenXml"%>
<%@ Import Namespace = "DocumentFormat.OpenXml.Packaging"%>
<%@ Import Namespace = "DocumentFormat.OpenXml.Wordprocessing"%>
<%@ Import Namespace = "A=DocumentFormat.OpenXml.Drawing" %>
<%@ Import Namespace = "DW=DocumentFormat.OpenXml.Drawing.Wordprocessing"%>
<%@ Import Namespace = "PIC=DocumentFormat.OpenXml.Drawing.Pictures"%>

<script runat="server">
	protected string in_scode = "";
	protected string in_no = "";
	protected string branch = "";
	protected string receipt_title = "";

	public IpoReport ipoRpt = null;
	protected string templateFile = "";
	protected string outputFile = "";

	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;

		in_scode = (Request["in_scode"] ?? "").ToString();//n100
		in_no = (Request["in_no"] ?? "").ToString();//20170103001
		branch = (Request["branch"] ?? "").ToString();//N
		receipt_title = (Request["receipt_title"] ?? "").ToString();//B

		DocxOutNewClass();
		//DocxImg();
		//DocxImg2();
		//DocxOut();
		//DocxOutInClass();
	}

	protected void DocxOutNewClass() {
		string templateFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx";
		ipoRpt = new IpoReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch, receipt_title);
		try {
			ipoRpt.CloneToStream(templateFile, true);
			
			DataTable dmp = ipoRpt.getDmp();
			if (dmp.Rows.Count > 0) {
				//標題區塊
				ipoRpt.CopyBlock("b_title");
				//一併申請實體審查
				if (dmp.Rows[0]["reality"].ToString() == "Y") {
					ipoRpt.ReplaceBookmark("reality", "是");
				} else {
					ipoRpt.ReplaceBookmark("reality", "否");
				}
				//事務所或申請人案件編號
				ipoRpt.ReplaceBookmark("seq", ipoRpt.getSeq() + "-" + dmp.Rows[0]["scode1"].ToString());
				//中文發明名稱 / 英文發明名稱
				ipoRpt.ReplaceBookmark("cappl_name", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
				ipoRpt.ReplaceBookmark("eappl_name", dmp.Rows[0]["eappl_name"].ToString().ToXmlUnicode());
				//申請人
				using (DataTable dtAp = ipoRpt.GetApcust()) {
					for (int i = 0; i < dtAp.Rows.Count; i++) {
						ipoRpt.CopyBlock("b_apply");
						ipoRpt.ReplaceBookmark("apply_num", (i + 1).ToString());
						ipoRpt.ReplaceBookmark("ap_country", dtAp.Rows[i]["Country_name"].ToString());
						ipoRpt.ReplaceBookmark("ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
						ipoRpt.ReplaceBookmark("ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
						ipoRpt.ReplaceBookmark("ap_cname", dtAp.Rows[i]["Cname_string"].ToString());
						ipoRpt.ReplaceBookmark("ap_ename", dtAp.Rows[i]["Ename_string"].ToString());
					}
				}
				//代理人
				ipoRpt.CopyBlock("b_agent");
				using (DataTable dtAgt = ipoRpt.GetAgent()) {
					ipoRpt.ReplaceBookmark("agt_name1", dtAgt.Rows[0]["agt_name1"].ToString().Trim());
					ipoRpt.ReplaceBookmark("agt_name2", dtAgt.Rows[0]["agt_name2"].ToString().Trim());
				}
				//發明人
				using (DataTable dtAnt = ipoRpt.GetAnt()) {
					for (int i = 0; i < dtAnt.Rows.Count; i++) {
						ipoRpt.CopyBlock("b_ant");
						ipoRpt.ReplaceBookmark("ant_num", "發明人" + (i + 1).ToString());
						ipoRpt.ReplaceBookmark("ant_country", dtAnt.Rows[i]["Country_name"].ToString());
						ipoRpt.ReplaceBookmark("ant_cname", dtAnt.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
						ipoRpt.ReplaceBookmark("ant_ename", dtAnt.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
					}
				}
				//主張優惠期
				ipoRpt.CopyBlock("b_exh");
				string exh_date = "";
				if (dmp.Rows[0]["exhibitor"].ToString() == "Y") {//參展或發表日期填入表中的發生日期
					if (dmp.Rows[0]["exh_date"] != System.DBNull.Value && dmp.Rows[0]["exh_date"] != null) {
						exh_date = Convert.ToDateTime(dmp.Rows[0]["exh_date"]).ToString("yyyy/MM/dd");
					}
				}
				ipoRpt.ReplaceBookmark("exh_date", exh_date);


				//主張利用生物材料/生物材料不須寄存/聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
				ipoRpt.CopyBlock("b_content");
				//聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
				if (dmp.Rows[0]["same_apply"].ToString() == "Y") {
					ipoRpt.ReplaceBookmark("same_apply", "是");
				} else {
					ipoRpt.ReplaceBookmark("same_apply", "");
				}
				//附送書件
				ipoRpt.CloneReplaceBlock("b_attach", "#seq#", ipoRpt.getSeq());
				ipoRpt.CopyBlock("b_sign");
				//docx.GenerateImageRun(Server.MapPath("~/ReportTemplate") + @"\66824.jpg");
				ipoRpt.CopyFoot();
				//docx.CopyNewPageFoot();
			}
			ipoRpt.Flush("-發明-" + DateTime.Now.ToString("yyyyMMdd") + ".docx");
		}
		catch (Exception ex) {
			throw ex;//Response.Write(ex.ToString());
		}
		finally {
			ipoRpt.Close();
		}
	}
	
	protected void DocxImg2() {
		templateFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx";
		outputFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書_out.docx";

		File.Copy(templateFile, outputFile,true);
		using (WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateFile, false)) {
			using (WordprocessingDocument outDoc = WordprocessingDocument.Open(outputFile, true)) {
				//抓頁尾
				SectionProperties foot = (SectionProperties)outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true);
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();

				Body body = outDoc.MainDocumentPart.Document.Body;
				body.Append(CopyBlock(tempDoc, "b_title"));
				ReplaceBookmark(outDoc.MainDocumentPart, "reality", "否");
				body.AppendChild(new Paragraph(GenerateImageRun(outDoc, new ImageData(Server.MapPath("~/ReportTemplate") + @"\66824.jpg"))));
			}
		}
		Response.End();
	}

	protected void DocxImg() {
		templateFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書_img.docx";
		outputFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書_out.docx";

		byte[] byteArray = File.ReadAllBytes(templateFile);
		using (WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateFile, false)) {
			//using (Stream Mem = File.Open(templateFile, FileMode.Open)) {
				using (MemoryStream Mem = new MemoryStream()) {
				Mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument outDoc = WordprocessingDocument.Open(Mem, true)) {
					//抓頁尾
					SectionProperties foot = (SectionProperties)outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true);
					outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
					outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
					outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();

					Body body = outDoc.MainDocumentPart.Document.Body;
					body.Append(CopyBlock(tempDoc, "b_title"));
					ReplaceBookmark(outDoc.MainDocumentPart, "reality", "否");
					body.Append(CopyBlock(tempDoc, "b_table"));					
					//body.AppendChild(new Paragraph(GenerateImageRun(outDoc, new ImageData(Server.MapPath("~/ReportTemplate") + @"\66824.jpg"))));
					//body.AppendChild(foot);//頁尾

					outDoc.MainDocumentPart.Document.Save();

					using (MemoryStream memoryStream = new MemoryStream()) {
						Mem.Position = 0;
						Mem.WriteTo(memoryStream);
						Response.Clear();
						Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5");
						Response.AddHeader("Content-Disposition", "attachment; filename=\"-新型-" + DateTime.Now.ToString("yyyyMMdd") + ".doc\"");
						Response.ContentType = "application/octet-stream";
						Response.OutputStream.Write(memoryStream.GetBuffer(), 0, memoryStream.GetBuffer().Length);
						//Mem.Position = 0;
						//Mem.CopyTo(Response.OutputStream);
						Response.OutputStream.Flush();
						Response.OutputStream.Close();
						Response.Flush();
					}
					/////////////////////////////
					//Byte[] byteArray1 = Mem.ToArray();
					//Mem.Flush();
					//Mem.Close();
					//Response.BufferOutput = true;
					//Response.Clear();
					//Response.ClearHeaders();
					//Response.AddHeader("Content-Disposition", "attachment; filename=\"-新型-" + DateTime.Now.ToString("yyyyMMdd") + ".doc\"");
					//Response.ContentType = "application/octet-stream";
					//// Write the data
					//Response.BinaryWrite(byteArray1);
					//Response.End();
					/////////////////////////////
					//using (FileStream file = new FileStream(outputFile, FileMode.Create, System.IO.FileAccess.Write)) {
					//	//byte[] bytes = new byte[Mem.Length];
					//	//Mem.Read(bytes, 0, (int)Mem.Length);
					//	file.Write(Mem.GetBuffer(), 0, (int)Mem.Length);
					//	//Mem.Close();
					//}
					/////////////////////////////
					// 把 Stream 轉換成 byte[]
					//byte[] bytes = new byte[Mem.Length];
					//Mem.Read(bytes, 0, bytes.Length);
					//// 設置當前流的位置為流的開始
					//Mem.Seek(0, SeekOrigin.Begin);
					//using (FileStream file = new FileStream(outputFile, FileMode.Create, System.IO.FileAccess.Write)) {
					//	//byte[] bytes = new byte[Mem.Length];
					//	//Mem.Read(bytes, 0, (int)Mem.Length);
					//	file.Write(bytes, 0, bytes.Length);
					//	//Mem.Close();
					//}
				}
			}
		}
		Response.End();
	}

	public static void CopyStream(Stream input, MemoryStream output) {
		byte[] buffer = new byte[16 * 1024];
		int read;
		while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
			output.Write(buffer, 0, read);
		}
	}

	protected void DocxOutInClass() {
		string templateFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx";
		ipoRpt = new IpoReport(Session["btbrtdb"].ToString(), in_scode, in_no, branch, receipt_title);
		try {
			ipoRpt.CloneToStream(templateFile,true);

			ipoRpt.CopyBlock("b_title");
			ipoRpt.ReplaceBookmark("reality", "否");
			ipoRpt.ReplaceBookmark("seq", "NP28758-n100");
			ipoRpt.ReplaceBookmark("cappl_name", "1112&峯3435");
			ipoRpt.ReplaceBookmark("eappl_name", "1112&FENG3435");
			ipoRpt.CopyBlock("b_apply");
			ipoRpt.ReplaceBookmark("apply_num", "1");
			ipoRpt.ReplaceBookmark("ap_country", "TW中華民國");
			ipoRpt.ReplaceBookmark("ap_cname_title", "中文名稱");
			ipoRpt.ReplaceBookmark("ap_ename_title", "英文名稱");
			ipoRpt.ReplaceBookmark("ap_cname", "英業達股份有限公司");
			ipoRpt.ReplaceBookmark("ap_ename", "INVENTEC&LIFE CORPORATION");
			ipoRpt.CopyBlock("b_agent");
			ipoRpt.ReplaceBookmark("agt_name1", "高,玉駿");
			ipoRpt.ReplaceBookmark("agt_name2", "楊,祺雄");
			ipoRpt.CopyBlock("b_ant");
			ipoRpt.ReplaceBookmark("ant_num", "發明人1");
			ipoRpt.ReplaceBookmark("ant_country", "AT奧地利");
			ipoRpt.ReplaceBookmark("ant_cname", "許,𥡪瑄𥡪");
			ipoRpt.ReplaceBookmark("ant_ename", "xu,yix&uang");
			ipoRpt.CopyBlock("b_exh");
			ipoRpt.ReplaceBookmark("exh_date", "");
			ipoRpt.CopyBlock("b_content");
			ipoRpt.ReplaceBookmark("same_apply", "");
			ipoRpt.CloneReplaceBlock("b_attach","#seq#", "NP28758");
			ipoRpt.CopyBlock( "b_sign");
			//docx.GenerateImageRun(Server.MapPath("~/ReportTemplate") + @"\66824.jpg");
			ipoRpt.CopyFoot();
			//docx.CopyNewPageFoot();

			ipoRpt.Flush("-發明-" + DateTime.Now.ToString("yyyyMMdd") + ".doc");
		}
		catch (Exception ex) {
			//Response.Write(ex.ToString());
		}
		finally {
			ipoRpt.Close();
		}
	}

	protected void DocxOut() {
		templateFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx";

		using (WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateFile, false)) {
			//byte[] byteArray = File.ReadAllBytes(templateFile);
			using (MemoryStream Mem = new MemoryStream(File.ReadAllBytes(templateFile))) {
				//Mem.Write(byteArray, 0, (int)byteArray.Length);
				using (WordprocessingDocument outDoc = WordprocessingDocument.Open(Mem, true)) {
					//抓頁尾
					SectionProperties foot = (SectionProperties)outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true);
					outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
					outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
					outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();

					Body body = outDoc.MainDocumentPart.Document.Body;
					body.Append(CopyBlock(tempDoc, "b_title"));
					ReplaceBookmark(outDoc.MainDocumentPart, "reality", "否");
					ReplaceBookmark(outDoc.MainDocumentPart, "seq", "NP28758-n100");
					ReplaceBookmark(outDoc.MainDocumentPart, "cappl_name", "1112&峯3435");
					ReplaceBookmark(outDoc.MainDocumentPart, "eappl_name", "1112&FENG3435");
					body.Append(CopyBlock(tempDoc, "b_apply"));
					ReplaceBookmark(outDoc.MainDocumentPart, "apply_num", "1");
					ReplaceBookmark(outDoc.MainDocumentPart, "ap_country", "TW中華民國");
					ReplaceBookmark(outDoc.MainDocumentPart, "ap_cname_title", "中文名稱");
					ReplaceBookmark(outDoc.MainDocumentPart, "ap_ename_title", "英文名稱");
					ReplaceBookmark(outDoc.MainDocumentPart, "ap_cname", "英業達股份有限公司");
					ReplaceBookmark(outDoc.MainDocumentPart, "ap_ename", "INVENTEC&LIFE CORPORATION");
					body.Append(CopyBlock(tempDoc, "b_agent"));
					ReplaceBookmark(outDoc.MainDocumentPart, "agt_name1", "高,玉駿");
					ReplaceBookmark(outDoc.MainDocumentPart, "agt_name2", "楊,祺雄");
					body.Append(CopyBlock(tempDoc, "b_ant"));
					ReplaceBookmark(outDoc.MainDocumentPart, "ant_num", "發明人1");
					ReplaceBookmark(outDoc.MainDocumentPart, "ant_country", "AT奧地利");
					ReplaceBookmark(outDoc.MainDocumentPart, "ant_cname", "許,𥡪瑄𥡪");
					ReplaceBookmark(outDoc.MainDocumentPart, "ant_ename", "xu,yix&uang");
					body.Append(CopyBlock(tempDoc, "b_exh"));
					ReplaceBookmark(outDoc.MainDocumentPart, "exh_date", "");
					body.Append(CopyBlock(tempDoc, "b_content"));
					ReplaceBookmark(outDoc.MainDocumentPart, "same_apply", "");
					body.Append(ReplaceBlock(CopyBlock(tempDoc, "b_attach"), "#seq#", "NP28758"));
					body.Append(CopyBlock(tempDoc, "b_sign"));
					//body.AppendChild(new Paragraph(GenerateImageRun(outDoc, new ImageData(Server.MapPath("~/ReportTemplate") + @"\66824.jpg"))));
					//body.Append(copyTagAndReplace(tempDoc, "b_title", new Dictionary<string, string>() { { "#case_no#", "112233" } }));

					//body.AppendChild(new Paragraph(new ParagraphProperties(foot)));//頁尾+換頁
					body.AppendChild(foot);//頁尾

					outDoc.MainDocumentPart.Document.Save();
					using (MemoryStream memoryStream = new MemoryStream()) {
						Mem.WriteTo(memoryStream);
						Response.Clear();
						Response.AddHeader("Content-Disposition", "attachment; filename=\"-新型-" + DateTime.Now.ToString("yyyyMMdd") + ".doc\"");
						Response.ContentType = "application/octet-stream";
						Response.OutputStream.Write(memoryStream.GetBuffer(), 0, memoryStream.GetBuffer().Length);
						Response.OutputStream.Flush();
						Response.OutputStream.Close();
						Response.Flush();
					}
				}
			}
		}
		Response.End();
	}

	private static Paragraph[] CopyBlock(WordprocessingDocument doc, string tagName) {
		List<Paragraph> arrElement = new List<Paragraph>();
		Tag elementTag = doc.MainDocumentPart.RootElement.Descendants<Tag>()
		.Where(
			element => element.Val.Value.ToLower() == tagName.ToLower()
		).SingleOrDefault();

		if (elementTag != null) {
			SdtElement block = (SdtElement)elementTag.Parent.Parent;
			IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
			foreach (Paragraph tagRun in tagRuns) {
				arrElement.Add((Paragraph)tagRun.CloneNode(true));
				//return tagRun.CloneNode(true);
			}
		}
		return arrElement.ToArray();
	}

	private static Paragraph[] ReplaceBlock(Paragraph[] pars, string searchStr, string newStr) {
		for (int i = 0; i < pars.Length; i++) {
			pars[i] = (new Paragraph(new Run(new Text(pars[i].InnerText.Replace(searchStr, newStr)))));
			//string newText = pars[i].InnerText.Replace(searchStr, newStr);
			//pars[i].RemoveAllChildren<Run>();
			//pars[i].Append(new Run(new Text(newText)));
			////pars[i] = pars[i].Append(new Text(newText));
		}
		return pars;
	}

	private static void ReplaceBookmark(MainDocumentPart documentPart, string bookmarkName, string text) {
		IEnumerable<BookmarkEnd> bookMarkEnds = documentPart.RootElement.Descendants<BookmarkEnd>();
		foreach (BookmarkStart bookmarkStart in documentPart.RootElement.Descendants<BookmarkStart>()) {
			if (bookmarkStart.Name.Value.ToLower() == bookmarkName.ToLower()) {
				string id = bookmarkStart.Id.Value;
				BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).First();

				//var bookmarkText = bookmarkEnd.NextSibling();
				Run bookmarkRun = bookmarkStart.NextSibling<Run>();
				if (bookmarkRun != null) {
					string[] txtArr = text.Split('\n');
					for (int i = 0; i < txtArr.Length; i++) {
						if (i == 0) {
							bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
						} else {
							bookmarkRun.Append(new Break());
							bookmarkRun.Append(new Text(txtArr[i]));
						}
					}
					//bookmarkRun.GetFirstChild<Text>().Text = text;
					//bookmarkRun.Append(new Break());
					//bookmarkRun.Append(new Text("換行"));
				}
				bookmarkStart.Remove();
				bookmarkEnd.Remove();
			}
		}
	}

	public static Run GenerateImageRun(WordprocessingDocument wordDoc, ImageData img) {
		MainDocumentPart mainPart = wordDoc.MainDocumentPart;

		ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
		var relationshipId = mainPart.GetIdOfPart(imagePart);
		imagePart.FeedData(img.getDataStream());

		// Define the reference of the image.
		var element =
			 new Drawing(
				 new DW.Inline(
					 //Size of image, unit = EMU(English Metric Unit)
					 //1 cm = 360000 EMUs
					 new DW.Extent() { Cx = img.GetWidthInEMU(), Cy = img.GetHeightInEMU() },
					 new DW.EffectExtent()
					 {
						 LeftEdge = 0L,
						 TopEdge = 0L,
						 RightEdge = 0L,
						 BottomEdge = 0L
					 },
					 new DW.DocProperties()
					 {
						 Id = (UInt32Value)1U,
						 Name = img.ImageName
					 },
					 new DW.NonVisualGraphicFrameDrawingProperties(
						 new A.GraphicFrameLocks() { NoChangeAspect = true }),
					 new A.Graphic(
						 new A.GraphicData(
							 new PIC.Picture(
								 new PIC.NonVisualPictureProperties(
									 new PIC.NonVisualDrawingProperties()
									 {
										 Id = (UInt32Value)0U,
										 Name = img.FileName
									 },
									 new PIC.NonVisualPictureDrawingProperties()),
								 new PIC.BlipFill(
									 new A.Blip(
										 new A.BlipExtensionList(
											 new A.BlipExtension()
											 {
												 Uri =
													"{28A0092B-C50C-407E-A947-70E740481C1C}"
											 })
									 )
									 {
										 Embed = relationshipId,
										 CompressionState =
										 A.BlipCompressionValues.Print
									 },
									 new A.Stretch(
										 new A.FillRectangle())),
								 new PIC.ShapeProperties(
									 new A.Transform2D(
										 new A.Offset() { X = 0L, Y = 0L },
										 new A.Extents()
										 {
											 Cx = img.GetWidthInEMU(),
											 Cy = img.GetHeightInEMU()
										 }),
									 new A.PresetGeometry(
										 new A.AdjustValueList()
									 )
									 { Preset = A.ShapeTypeValues.Rectangle }))
						 )
						 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
				 )
				 {
					 DistanceFromTop = (UInt32Value)0U,
					 DistanceFromBottom = (UInt32Value)0U,
					 DistanceFromLeft = (UInt32Value)0U,
					 DistanceFromRight = (UInt32Value)0U,
					 // EditId = "50D07946"
				 });
		return new Run(element);
	}

</script>
