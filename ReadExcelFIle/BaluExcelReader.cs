using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using Microsoft.Office.Interop.Excel;
//using Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelWord = Microsoft.Office.Interop.Word;
//using DocumentFormat.OpenXml.Spreadsheet;
//using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
//using PDF = iTextSharp;
//using iTextSharp.text;
//using iTextSharp.text.pdf;
//using iTextSharp.text.html.simpleparser;
//using EvoPdf;
using SelectPdf;

namespace BaluExcelToPDF
{
    public partial class BaluExcelReader : Form
    {
        public BaluExcelReader()
        {
            InitializeComponent();
        }

        DataSet result;
        private void btnOpen_Click(object sender, EventArgs e)
        {
            #region Excel
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            Dictionary<string,string> ListOfSheets = new Dictionary<string,string>();

            double dbr;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;
            string OutPutPdfName = string.Empty;
            StringBuilder SB = new StringBuilder();
            
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(ofd.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    OutPutPdfName = ofd.SafeFileName.Split('.')[0];

                    for (int i = 1; i <= xlWorkBook.Worksheets.Count; i++)
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                        range = xlWorkSheet.UsedRange;
                        rw = range.Rows.Count;
                        cl = range.Columns.Count;
                        SB = new StringBuilder();
                        SB.Append("<html><head>Balu Works -- "+xlWorkSheet.Name+ "</head><body><table style='border: 1px solid #1C6EA4;background-color: #EEEEEE;width: 100%;text-align: left;border-collapse: collapse;'>");
                        for (rCnt = 1; rCnt <= rw; rCnt++)
                        {
                            SB.Append("<tr>");
                            for (cCnt = 1; cCnt <= cl; cCnt++)
                            {
                                string str = ""+(range.Cells[rCnt, cCnt] as Excel.Range).Value;
                                //MessageBox.Show(str);
                                SB.Append("<td style= 'border: 1px solid #AAAAAA;padding: 3px 2px;'>");
                                SB.Append(str);
                                SB.Append("</td>");
                            }
                            SB.Append("</tr>");
                        }
                        SB.Append("</body></html>");
                        ListOfSheets.Add(xlWorkSheet.Name, SB.ToString());
                    }

                    xlWorkBook.Close(true, null, null);
                    xlApp.Quit();
                }
            }
            #endregion

            #region SelectPDF
            foreach(KeyValuePair<string, string> OneOnOnePdf in ListOfSheets)
            {
                HtmlToPdf converter = new HtmlToPdf();
                converter.Options.PdfPageSize = PdfPageSize.A4;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.WebPageWidth = 1024;
                converter.Options.WebPageHeight = 0;
                string Date = DateTime.Now.ToShortDateString();
                SelectPdf.PdfDocument doc = converter.ConvertHtmlString(OneOnOnePdf.Value, string.Empty);
                string path = @"D:\PDFSamples\" + OutPutPdfName + "-"+ OneOnOnePdf.Key + "--" + Date + "--" + Guid.NewGuid().ToString() + ".pdf";
                doc.Save(path);
                doc.Close();
            }
            #endregion

            MessageBox.Show("Your Balu has prepared File(s). Rock on :)", "Congragulations");
            #region WordRegion

            //Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            //app.Visible = true;
            //ExcelWord.Document doc = app.Documents.Add();
            //ExcelWord.Range rng = app.ActiveDocument.Range(0, 0);
            //rng.Text = SB.ToString();
            //object fileName = "D:\\BaluChat.docx";
            //object Missing = Type.Missing;

            //ExcelWord.ContentControl contentControl = doc.ContentControls.Add(ExcelWord.WdContentControlType.wdContentControlRichText,ref Missing);
            //contentControl.Title = "This is Balu Word";

            //string html = "";
            //string htmlTempFilePath = Path.Combine(Path.GetTempPath(), string.Format("{0}.html", Path.GetRandomFileName()));

            //using (StreamWriter writer = File.CreateText(htmlTempFilePath))
            //{
            //    html = string.Format("{0}", SB.ToString());
            //    writer.WriteLine(html);
            //}

            //contentControl.Range.InsertFile(html,ref Missing, ref Missing, ref Missing, ref Missing);

            //doc = app.Documents.Open(fileName, Missing, Missing);
            //app.Selection.Find.ClearFormatting();
            //app.Selection.Find.Replacement.ClearFormatting();



            // app.Selection.HTMLDivisions.Add(SB);

            #endregion
        }

        
    }
}
