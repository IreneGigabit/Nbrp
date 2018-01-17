using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Drawing;

/// <summary>
/// Docx 操作類別(use OpenXml SDK)
/// </summary>
public class OpenXmlHelper {
	protected string _templateFile { get; set; }
	//protected SectionProperties Footer1 = null;//申請書頁尾
	//protected SectionProperties Footer2 = null;//基本資料表頁尾
	protected SectionProperties[] footer = null;
	protected WordprocessingDocument tempDoc = null;
	protected WordprocessingDocument outDoc = null;
	protected MemoryStream tempMem = null;
	protected MemoryStream outMem = new MemoryStream();
	protected Body outBody = null;

	public OpenXmlHelper() {
	}

	/// <summary>
	/// 關閉
	/// </summary>
	protected void Close() {
		if (this.tempDoc != null) tempDoc.Close();
		if (this.outDoc != null) outDoc.Close();
		if (this.tempMem != null) tempMem.Close();
		if (this.outMem != null) outMem.Close();
		HttpContext.Current.Response.End();
	}

	/// <summary>
	/// 複製範本檔到MemoryStream
	/// </summary>
	/// <param name="templateFile">範本檔名(實體路徑)</param>
	/// <param name="cleanFlag">是否清空內容(只保留版面配置)</param>
	public void CloneToStream(string templateFile,bool cleanFlag) {
		_templateFile = templateFile;

		tempDoc = WordprocessingDocument.Open(_templateFile, false);
		tempMem = new MemoryStream(File.ReadAllBytes(_templateFile));
		outDoc = WordprocessingDocument.Open(tempMem, true);

		footer = outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().ToArray();
		//抓取頁尾
		//Footer1 = (SectionProperties)outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true);
		//Footer2 = (SectionProperties)outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().LastOrDefault().CloneNode(true);

		//清空內容
		if (cleanFlag) {
			outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
			outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
			outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();
		}

		outBody = outDoc.MainDocumentPart.Document.Body;
	}

	/// <summary>
	/// 輸出檔案
	/// </summary>
	public void Flush(string outputName) {
		outDoc.MainDocumentPart.Document.Save();
		tempMem.WriteTo(outMem);
		HttpContext.Current.Response.Clear();
		HttpContext.Current.Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5");
		HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + outputName + "\"");
		HttpContext.Current.Response.ContentType = "application/octet-stream";
		HttpContext.Current.Response.OutputStream.Write(outMem.GetBuffer(), 0, outMem.GetBuffer().Length);
		HttpContext.Current.Response.OutputStream.Flush();
		HttpContext.Current.Response.OutputStream.Close();
		HttpContext.Current.Response.Flush();
		this.Close();
	}

	/// <summary>
	/// 增加段落文字
	/// </summary>
	public void AddParagraph(string text) {
		outDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(text))));
	}

	/// <summary>
	/// 複製範本Block
	/// </summary>
	public void CopyBlock(string blockName) {
		Tag elementTag = tempDoc.MainDocumentPart.RootElement.Descendants<Tag>()
		.Where(
			element => element.Val.Value.ToLower() == blockName.ToLower()
		).SingleOrDefault();

		if (elementTag != null) {
			SdtElement block = (SdtElement)elementTag.Parent.Parent;
			IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
			foreach (Paragraph tagRun in tagRuns) {
				//arrElement.Add((Paragraph)tagRun.CloneNode(true));
				outBody.Append((Paragraph)tagRun.CloneNode(true));
			}
		}
	}

	/// <summary>
	/// 複製範本Block,回傳Array
	/// </summary>
	private Paragraph[] CopyBlockArry(string blockName) {
		List<Paragraph> arrElement = new List<Paragraph>();
		Tag elementTag = tempDoc.MainDocumentPart.RootElement.Descendants<Tag>()
		.Where(
			element => element.Val.Value.ToLower() == blockName.ToLower()
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

	/// <summary>
	/// 取代書籤
	/// </summary>
	public void ReplaceBookmark(string bookmarkName, string text) {
		try {
			MainDocumentPart mainPart = outDoc.MainDocumentPart;
			IEnumerable<BookmarkEnd> bookMarkEnds = mainPart.RootElement.Descendants<BookmarkEnd>();
			foreach (BookmarkStart bookmarkStart in mainPart.RootElement.Descendants<BookmarkStart>()) {
				if (bookmarkStart.Name.Value.ToLower() == bookmarkName.ToLower()) {
					string id = bookmarkStart.Id.Value;
					//BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).First();
					BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).FirstOrDefault();

					////var bookmarkText = bookmarkEnd.NextSibling();
					//Run bookmarkRun = bookmarkStart.NextSibling<Run>();
					//if (bookmarkRun != null) {
					//	string[] txtArr = text.Split('\n');
					//	for (int i = 0; i < txtArr.Length; i++) {
					//		if (i == 0) {
					//			bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
					//		} else {
					//			bookmarkRun.Append(new Break());
					//			bookmarkRun.Append(new Text(txtArr[i]));
					//		}
					//	}
					//}
					Run bookmarkRun = bookmarkStart.NextSibling<Run>();
					if (bookmarkRun != null) {
						Run tempRun = bookmarkRun;
						string[] txtArr = text.Split('\n');
						for (int i = 0; i < txtArr.Length; i++) {
							if (i == 0) {
								bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
							} else {
								bookmarkRun.Append(new Break());
								bookmarkRun.Append(new Text(txtArr[i]));
							}
						}
						int j=0;
						while (tempRun.NextSibling()!=null && tempRun.NextSibling().GetType() != typeof(BookmarkEnd)) {
							j++;
							tempRun.NextSibling().Remove();
							if (j >= 20)
								break;
						}
					}
					bookmarkStart.Remove();
					if (bookmarkEnd!=null)bookmarkEnd.Remove();
				}
			}
		}
		catch (Exception ex) {
			throw new Exception("取代書籤錯誤!!(" + bookmarkName + ")", ex);
		}
	}

	/// <summary>
	/// 複製範本Block,並取代文字
	/// </summary>
	public void CloneReplaceBlock(string blockName, string searchStr, string newStr) {
		Paragraph[] pars = CopyBlockArry(blockName);
		for (int i = 0; i < pars.Length; i++) {
			pars[i] = (new Paragraph(new Run(new Text(pars[i].InnerText.Replace(searchStr, newStr)))));
		}
		outBody.Append(pars);
	}

	/// <summary>
	/// 複製頁尾,且增加新頁
	/// </summary>
	public void AppendNewPageFoot(int index) {
		//outBody.AppendChild(new Paragraph(new ParagraphProperties(Footer1)));//頁尾+換頁
		outBody.AppendChild(new Paragraph(new ParagraphProperties(footer[index].CloneNode(true))));//頁尾+換頁
	}

	/// <summary>
	/// 複製頁尾
	/// </summary>
	public void AppendFoot(int index) {
		outBody.AppendChild(footer[index].CloneNode(true));
	}
}