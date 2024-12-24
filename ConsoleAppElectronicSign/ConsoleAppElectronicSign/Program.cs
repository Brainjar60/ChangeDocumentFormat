
using Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word.Document;
namespace ConsoleAppElectronicSign
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            SaveAsPDF(); 
        }
        public static string ChgFileExt(string fileName, string newExt)
        {
            return fileName.Substring(0, fileName.LastIndexOf(".") + 1) + newExt;
        }
        public static void SaveAsPDF()
        {
            string newDocument = "C:/ProveRepos/DocumentoDaFirmareConSignHereAnchorPagina1.docx";
            // WriteFile(byteArray, newDocument);
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null;
            application.Visible = false;
            // open docx file
            document = application.Documents.Open(newDocument);
            // convert to pdf file
            document.ExportAsFixedFormat(ChgFileExt(newDocument, "pdf"), WdExportFormat.wdExportFormatPDF);
            document.Close();
        }
    }
}
