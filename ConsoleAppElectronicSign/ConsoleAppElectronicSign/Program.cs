
using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Diagnostics; 
using System.Net.Mail;
using System.Reflection.Metadata;
using Document = Microsoft.Office.Interop.Word.Document;
using MailMessage = System.Net.Mail.MailMessage;
namespace ConsoleAppElectronicSign
{
    public class Program
    {
        private static string attDocument = "";

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            SaveAsPDF();
            SendMail(); 
        }
        public static string ChgFileExt(string fileName, string newExt)
        {
            return fileName.Substring(0, fileName.LastIndexOf(".") + 1) + newExt;
        }
        public static void SaveAsPDF()
        {
            string newDocument = "E:/ProveRepos/DocumentoDaFirmareConSignHereAnchorPagina1.docx";
            // WriteFile(byteArray, newDocument);
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null;
            application.Visible = false;
            // open docx file
            document = application.Documents.Open(newDocument);
            // convert to pdf file
            attDocument = ChgFileExt(newDocument, "pdf");
            document.ExportAsFixedFormat(attDocument, WdExportFormat.wdExportFormatPDF);
            document.Close();
            application.Quit();
            application = null; 
        }
        public static void SendMail()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.atvinformatica.it", 25);
                SmtpServer.UseDefaultCredentials = false;
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.OnFailure;
                SmtpServer.Credentials = new System.Net.NetworkCredential("postmaster@atvinformatica.it", "ArubaPassword$0");
                SmtpServer.EnableSsl = false;
                mail.From = new MailAddress("postmaster@atvinformatica.it");
                ArrayList emailingList = new ArrayList();
                emailingList.Add("vito.perrotta@ipsedocet.it");
                emailingList.Add("v.perrotta@dacomat.com");
                emailingList.Add("vito.perrotta@gmail.com");
                foreach (string oTo in emailingList)
                {
                    mail.To.Add(oTo);
                }
                mail.CC.Add("vperrotta@alveria.it");
                mail.Attachments.Add(new Attachment(attDocument));
                mail.Subject = "Invio programma SMT";
                mail.Body = "messaggio";
                SmtpServer.Send(mail);
                Debug.WriteLine("Mail spedita correttamente");
                Console.WriteLine("Mail spedita correttamente");

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Console.WriteLine(ex.Message);
            }
        }
        /// <summary>
        /// Metodo da inserire in un controller ASPX
        /// </summary>
        /// <param name="envelopeId"></param>
        /// <param name="documentId"></param>
        /// <param name="documentName"></param>
        /// <returns></returns>
        public string GetEnvelopeDocument(string envelopeId, string documentId, string documentName)
        {
            Stream streamVal = null; // Acuisire lo stream del documento
            string fileName = "c:/Temp/" + documentName + "_" + envelopeId + ".pdf";
            //saveTemporaryFile(fileName, streamVal); 
            int iB = (int)streamVal.Length;
            BinaryReader reader = new BinaryReader(streamVal);
            byte[] pdfBytes = reader.ReadBytes(iB);
            string docBase64 = "" + Convert.ToBase64String(pdfBytes);
            return docBase64;
        }
        // ------------------------------------------------------------------------------------
        // funzione Javascript che legge il documento tramite metodo controller e lo visualizza
        // in una finestra della applicazione host 
        // ------------------------------------------------------------------------------------
        /*
        function invokeAjaxGetEnvelopeCertificate(envelopeId, documentId, documentName)
        {
            var xhr = new XMLHttpRequest();
            var _url = 'api/adlinker/ajaxGetEnvelopeDocument?envelopeId=' + envelopeId + "&documentId=" + documentId + "&documentName=" + documentName;
            xhr.onreadystatechange = function() {
                if (this.readyState == 4 && this.status == 200)
                {
                    if (!this.response || this.response == null || this.response == "")
                    {
                        return;
                    }
                    console.log("ajax response" + documentName);

                    let pdfWindow = window.open("Certificate of Completion", "Certificate", "titlebar=yes,toolbar=yes,scrollbars=yes,resizable=yes,top=10,left=10,width=1725,height=865");
                    pdfWindow.document.write(
                        "<iframe width='100%' height='100%' src='data:application/pdf;base64, " +
                        this.response + "'></iframe>"
                    )

                }
                else
                {
                    return;
                }
            };
            xhr.open('get', _url);
            xhr.send();
        }
        */
    }
}
