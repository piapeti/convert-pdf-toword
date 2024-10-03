using System;
using Microsoft.Office.Interop.Word;
using Spire.Pdf;
namespace ConsoleApp5
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Application wordApp = new Application();
            //Document wordDocument = wordApp.Documents.Open(@"D:\PriyaUpkare_Immediate_Joiner.docx");

            //wordDocument.ExportAsFixedFormat(@"E:\documentsnew.pdf", WdExportFormat.wdExportFormatPDF);

            //wordDocument.Close();
            //wordApp.Quit();



            //convert pdf to word
            convertpdf();
        }
    static    void convertpdf() {

            PdfDocument pdfDocument = new PdfDocument();
            pdfDocument.LoadFromFile(@"E:\documentsnew.pdf");

            pdfDocument.SaveToFile(@"E:\path123.docx", FileFormat.DOCX);
        }
    }
}
