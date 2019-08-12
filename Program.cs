using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace PDF_to_DOCX_via_NuGet
{
    class Program
    {
        static void Main(string[] args)
        {

            DirectoryInfo di = new DirectoryInfo(args[0]);
            var files = di.GetFiles();
            Application app = new Application();
            foreach (FileInfo fi in files)
            {
                if (fi.FullName.EndsWith(".pdf"))
                {
                    // Convert PDF file to DOCX file 
                    //SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
                    Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(fi.FullName);
                    //f.OpenPdf(fi.FullName);
                    if (doc != null)//.PageCount > 0)
                    {
                        Console.WriteLine($"Converting file {fi.FullName}");
                        // You may choose output format between Docx and Rtf. 
                        //f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;
                        string fullName = fi.FullName.Replace(".pdf", ".docx");
                        //int result = f.ToWord(fullName);
                        //var document = app.Documents.Open(fullName);
                        doc.SaveAs2(fullName, WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: WdCompatibilityMode.wdCurrent);
                        app.ActiveDocument.Close();
                    }
                }
            }
            app.Quit();
        }
    }
}