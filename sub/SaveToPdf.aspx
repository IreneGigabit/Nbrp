<%@ Page Language="C#" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Collections.Generic"%>
<%@ Import Namespace = "Word=Microsoft.Office.Interop.Word" %>

<script runat="server">
    object srcFile = "";
    object destFils = "";

    //word用的常數值==
    object wdFormatPDF = Word.WdSaveFormat.wdFormatPDF;
    object oFalse = false;
    object oTrue = true;
	object oMissing = System.Reflection.Missing.Value;
    //===============
    
	protected Word._Application wordApp = null;
    //http://web08/nbrp/sub/SaveToPdf.aspx
    private void Page_Load(System.Object sender, System.EventArgs e) {
        srcFile = (Request["srcFile"] ?? "").ToString();
        destFils = (Request["destFils"] ?? "").ToString();
        
        srcFile = @"\\web02\brp\reportdata\letter_dmp_NP-1090-M2007_br.xml";
        destFils = @"\\web02\brp\reportdata\letter_dmp_NP-1090-M2007_br.pdf";

        wordApp = new Word.Application();

        object oFilePath = srcFile;    //檔案路徑
        Word._Document myDoc = wordApp.Documents.Open(ref oFilePath, ref oMissing, ref oTrue, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oFalse,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        myDoc.Activate();
        try {
            wordApp.ActiveDocument.SaveAs(ref destFils, ref wdFormatPDF, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }
        catch (Exception ex) {
            Response.Write("<Font align=left color=\"red\" size=3>Eeception - " + ex.Message + "!!</font><BR>");
        }
        finally {
            wordApp.ActiveDocument.Close(ref oMissing, ref oMissing, ref oMissing);
            wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
            if (myDoc != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myDoc);
            if (wordApp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            myDoc = null;
            wordApp = null;
            GC.Collect();
        }
    }
</script>
