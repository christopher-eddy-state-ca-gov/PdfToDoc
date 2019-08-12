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
            string wd = ".";
            if(args.Length > 0)
            {
                wd = args[0];
            }
            DirectoryInfo di = new DirectoryInfo(wd);
            var files = di.GetFiles();
            Application app = new Application();
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            foreach (FileInfo fi in files)
            {
                if (fi.FullName.EndsWith(".pdf"))
                {
                    // Convert PDF file to DOCX file 
                    Microsoft.Office.Interop.Word.Document f = app.Documents.Open(fi.FullName);
                    if (f != null) 
                    {
                        Console.WriteLine($"Converting file {fi.FullName}");
                        string fullName = fi.FullName.Replace(".pdf", ".docx");
                        f.SaveAs2(fullName, WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: WdCompatibilityMode.wdCurrent);
                        app.ActiveDocument.Close();
                    }
                }
            }
            app.Quit();
        }
    }
}