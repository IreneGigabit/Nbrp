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
	protected WordprocessingDocument outDoc = null;
	protected MemoryStream outMem = new MemoryStream();
	protected Body outBody = null;
	Dictionary<string, WordprocessingDocument> tplDoc = new Dictionary<string, WordprocessingDocument>();
	protected string defTplDocName = "";
	Dictionary<string, MemoryStream> tplMem = new Dictionary<string, MemoryStream>();

	public OpenXmlHelper() {
	}

	#region 關閉
	/// <summary>
	/// 關閉
	/// </summary>
	public void Dispose() {
		if (this.outDoc != null) outDoc.Dispose();
		if (this.outMem != null) outMem.Close();

		foreach (KeyValuePair<string, WordprocessingDocument> item in tplDoc) {
			item.Value.Dispose();
		}
		foreach (KeyValuePair<string, MemoryStream> item in tplMem) {
			item.Value.Close();
			item.Value.Dispose();
		}
	}
	#endregion

	#region 建立空白檔案
	/// <summary>
	/// 建立空白檔案
	/// </summary>
	public void Create() {
		MemoryStream outMem = new MemoryStream();
		outDoc = WordprocessingDocument.Create(outMem, WordprocessingDocumentType.Document);
		MainDocumentPart mainPart = outDoc.AddMainDocumentPart();
		mainPart.Document = new Document();
		outBody = mainPart.Document.AppendChild(new Body());
	}
	#endregion

	#region 複製範本檔
	/// <summary>
	/// 複製範本檔
	/// </summary>
	/// <param name="templateList">範本＜別名,檔名(實體路徑)＞</param>
	public void CloneFromFile(Dictionary<string, string> templateList, bool cleanFlag) {
		foreach (var x in templateList.Select((Entry, Index) => new { Entry, Index })) {
			if (x.Index == 0) {
				byte[] outArray = File.ReadAllBytes(x.Entry.Value);
				outMem.Write(outArray, 0, (int)outArray.Length);
				outDoc = WordprocessingDocument.Open(outMem, true);
				defTplDocName = x.Entry.Key;
			}

			byte[] tplArray = File.ReadAllBytes(x.Entry.Value);
			tplMem.Add(x.Entry.Key, new MemoryStream());
			tplMem[x.Entry.Key].Write(tplArray, 0, (int)tplArray.Length);
			tplDoc.Add(x.Entry.Key, WordprocessingDocument.Open(tplMem[x.Entry.Key], false));
		}

		//清空輸出檔內容
		if (cleanFlag) {
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();
			outDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
		}

		outBody = outDoc.MainDocumentPart.Document.Body;
	}
	#endregion

	#region 輸出檔案
	/// <summary>
	/// 輸出檔案
	/// </summary>
	public void Flush(string outputName) {
		outDoc.MainDocumentPart.Document.Save();
		outDoc.Close();
		byte[] byteArray = outMem.ToArray();
		HttpContext.Current.Response.Clear();
		HttpContext.Current.Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5");
		HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + outputName + "\"");
		HttpContext.Current.Response.ContentType = "application/octet-stream";
		HttpContext.Current.Response.AddHeader("Content-Length", outMem.Length.ToString());
		HttpContext.Current.Response.BinaryWrite(outMem.ToArray());
		//微軟KB 312629https://support.microsoft.com/en-us/help/312629/prb-threadabortexception-occurs-if-you-use-response-end--response-redi
		///Response.End、Server.Transfer、Response.Redirect被呼叫時，會觸發ThreadAbortException，因此要改用CompleteRequest()
		//HttpContext.Current.Response.End();
		HttpContext.Current.ApplicationInstance.CompleteRequest();
		this.Dispose();
	}
	#endregion

	#region 增加段落文字
	/// <summary>
	/// 增加段落文字
	/// </summary>
	public void AddParagraph(string text) {
		outDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(text))));
	}
	#endregion

	#region 複製範本Block
	/// <summary>
	/// 複製範本Block
	/// </summary>
	public void CopyBlock(string blockName) {
		CopyBlock(defTplDocName, blockName);
	}

	/// <summary>
	/// 複製範本Block(指定文件)
	/// </summary>
	public void CopyBlock(string srcDocName, string blockName) {
		WordprocessingDocument srcDoc = tplDoc[srcDocName];
		Tag elementTag = srcDoc.MainDocumentPart.RootElement.Descendants<Tag>()
		.Where(
			element => element.Val.ToString().ToLower() == blockName.ToLower()
		).SingleOrDefault();

		if (elementTag != null) {
			SdtElement block = (SdtElement)elementTag.Parent.Parent;
			IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
			foreach (Paragraph tagRun in tagRuns) {
				outBody.Append((OpenXmlElement)tagRun.CloneNode(true));
			}
		}
	}
	#endregion

	#region 複製範本Block,回傳List
	/// <summary>
	/// 複製範本Block,回傳List
	/// </summary>
	public List<Paragraph> CopyBlockList(string blockName) {
		return CopyBlockList(defTplDocName, blockName);
	}

	/// <summary>
	/// 複製範本Block,回傳List(指定文件)
	/// </summary>
	private List<Paragraph> CopyBlockList(string srcDocName,string blockName) {
		try {
			WordprocessingDocument srcDoc = tplDoc[srcDocName]; 
			List<Paragraph> arrElement = new List<Paragraph>();
			Tag elementTag = srcDoc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower() == blockName.ToLower()
			).SingleOrDefault();
	
			if (elementTag != null) {
				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				IEnumerable<Paragraph> tagPars = block.Descendants<Paragraph>();
				foreach (Paragraph tagPar in tagPars) {
					arrElement.Add((Paragraph)tagPar.CloneNode(true));
				}
			}
			return arrElement;
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block!!(" + blockName + ")", ex);
		}
	}
	#endregion

	#region 複製範本Block,並取代文字
	/// <summary>
	/// 複製範本Block,並取代文字
	/// </summary>
	public void CloneReplaceBlock(string blockName, string searchStr, string newStr) {
		CloneReplaceBlock(defTplDocName, blockName, searchStr, newStr);
	}

	/// <summary>
	/// 複製範本Block,並取代文字(指定文件)
	/// </summary>
	public void CloneReplaceBlock(string srcDocName, string blockName, string searchStr, string newStr) {
		try {
			List<Paragraph> pars = CopyBlockList(srcDocName, blockName);
			for (int i = 0; i < pars.Count; i++) {
				pars[i] = (new Paragraph(new Run(new Text(pars[i].InnerText.Replace(searchStr, newStr)))));
			}
			outBody.Append(pars.ToArray());
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block錯誤!!(" + blockName + ")", ex);
		}
	}
	#endregion

	#region 取代書籤
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

					Run bookmarkRun = bookmarkStart.NextSibling<Run>();
					if (bookmarkRun != null) {
						Run tplRun = bookmarkRun;
						string[] txtArr = text.Split('\n');
						for (int i = 0; i < txtArr.Length; i++) {
							if (i == 0) {
								bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
							} else {
								bookmarkRun.Append(new Break());
								bookmarkRun.Append(new Text(txtArr[i]));
							}
						}
						int j = 0;
						while (tplRun.NextSibling() != null && tplRun.NextSibling().GetType() != typeof(BookmarkEnd)) {
							j++;
							tplRun.NextSibling().Remove();
							if (j >= 20)
								break;
						}
					}
					bookmarkStart.Remove();
					if (bookmarkEnd != null) bookmarkEnd.Remove();
				}
			}
		}
		catch (Exception ex) {
			throw new Exception("取代書籤錯誤!!(" + bookmarkName + ")", ex);
		}
	}
	#endregion

	#region 複製範本頁尾
	/// <summary>
	/// 複製範本頁尾
	/// </summary>
	/// <param name="sourceDoc">複製來源</param>
	/// <param name="haveBreak">是否帶分節符號(新頁)</param>
	public void CopyPageFoot(string srcDocName, bool haveBreak) {
		WordprocessingDocument sourceDoc = tplDoc[srcDocName];
		int index = 0;//取消index參數,只抓第1個

		string newRefId = string.Format("foot_{0}", Guid.NewGuid().ToString().Substring(0, 8));

		FooterReference[] footerSections = sourceDoc.MainDocumentPart.RootElement.Descendants<FooterReference>().ToArray();
		string srcRefId = footerSections[index].Id;
		footerSections[index].Id = newRefId;

		FooterPart elementFoot = sourceDoc.MainDocumentPart.FooterParts
		.Where(
			element => sourceDoc.MainDocumentPart.GetIdOfPart(element) == srcRefId
		).SingleOrDefault();
		outDoc.MainDocumentPart.AddPart(elementFoot, newRefId);

		if (haveBreak)
			outBody.AppendChild(new Paragraph(new ParagraphProperties(footerSections[index].Parent.CloneNode(true))));//頁尾+分節符號
		else
			outBody.AppendChild(footerSections[index].Parent.CloneNode(true));//頁尾
	}
	#endregion

	#region 插入圖片
	/// <summary>
	/// 插入圖片
	/// </summary>
	//public void AppendImage(string imgStr, bool isBase64, decimal scale) {
	//	ImageData img= new ImageData(imgStr, isBase64, scale);
	public void AppendImage(ImageFile img) {
		ImagePart imagePart = outDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
		string relationshipId = outDoc.MainDocumentPart.GetIdOfPart(imagePart);
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
									 ) { Preset = A.ShapeTypeValues.Rectangle }))
						 ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
				 )
				 {
					 DistanceFromTop = (UInt32Value)0U,
					 DistanceFromBottom = (UInt32Value)0U,
					 DistanceFromLeft = (UInt32Value)0U,
					 DistanceFromRight = (UInt32Value)0U,
					 //EditId = "50D07946"
				 });

		outDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
	}
	#endregion
}

public class ImageFile {
	public string FileName = string.Empty;

	public byte[] BinaryData;

	public Stream getDataStream() {
		//Stream DataStream = new MemoryStream(BinaryData);
		return new MemoryStream(BinaryData);
	}

	public ImagePartType ImageType {
		get {
			var ext = Path.GetExtension(FileName).TrimStart('.').ToLower();
			switch (ext) {
				case "jpg":
					return ImagePartType.Jpeg;
				case "png":
					return ImagePartType.Png;
				case "bmp":
					return ImagePartType.Bmp;
			}
			throw new ApplicationException(string.Format("不支援的格式:{0}", ext));
		}
	}

	public int SourceWidth;
	public int SourceHeight;
	public decimal Width;
	public decimal Height;

	//public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);
	private long WidthInEMU = 0;
	public long GetWidthInEMU() {
		WidthInEMU = Convert.ToInt64(Width * CM_TO_EMU);
		return WidthInEMU;
	}

	//public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);
	private long HeightInEMU = 0;
	public long GetHeightInEMU() {
		HeightInEMU = Convert.ToInt64(Height * CM_TO_EMU);
		return HeightInEMU;
	}

	private const decimal INCH_TO_CM = 2.54M;
	private const decimal CM_TO_EMU = 360000M;
	public string ImageName;

	public ImageFile(string fileName, byte[] data, decimal scale) {
		if (fileName == "") {
			FileName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
			ImageName = FileName;
		} else {
			FileName = fileName;
			ImageName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
		}

		BinaryData = data;
		int dpi = 300;
		Bitmap img = new Bitmap(new MemoryStream(data));
		SourceWidth = img.Width;
		SourceHeight = img.Height;
		Width = ((decimal)SourceWidth) / dpi * scale * INCH_TO_CM;
		Height = ((decimal)SourceHeight) / dpi * scale * INCH_TO_CM;
	}

	public ImageFile(byte[] data) :
		this("", data, 1) {
	}

	public ImageFile(byte[] data, decimal scale) :
		this("", data, scale) {
	}

	public ImageFile(string fileName) :
		this(fileName, File.ReadAllBytes(fileName), 1) {
	}

	public ImageFile(string fileName, decimal scale) :
		this(fileName, File.ReadAllBytes(fileName), scale) {
	}
}