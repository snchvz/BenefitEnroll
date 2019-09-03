using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NewEnrollmentsProgram
{
    /// <summary>
    /// Interaction logic for MergePage.xaml
    /// </summary>
    public partial class MergePage : Page
    {
        public MergePage()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            Object oTemplatePath = @"D:\Documents\PerformanceMergeTemplate.doc";
            try
            {
                oWordDoc = oWord.Documents.Open(oTemplatePath);
            }

            catch
            {
                MessageBox.Show("could not find performance temple doc");
                oWordDoc.Close(0);
                oWord.Quit();
                return;
            }

            //delete previous Performance review Doc
            DirectoryInfo dir = new DirectoryInfo(@"D:\Documents");
            foreach (FileInfo file in dir.GetFiles())
            {
                if (file.ToString().Contains("EMPLOYEE PERFORMANCE REVIEW.docx"))
                    file.Delete();
            }

            Object oMissing = System.Reflection.Missing.Value;
            object qry = "select * from [Sheet1$]";

            try
            {
                oWordDoc.MailMerge.OpenDataSource(@"D:\Desktop\TestFolder\TestExcel.xlsx", ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref qry, ref oMissing,
                ref oMissing, ref oMissing);
            }

            catch
            {
                MessageBox.Show("error opening excel data source");
                oWordDoc.Close(0);
                oWord.Quit();
                return;
            }

            oWordDoc.MailMerge.Execute(ref oMissing);

            //oWordDoc.SaveAs2(@"D:\Documents\MergeTemplate.doc");

            //oWordDoc.ExportAsFixedFormat("D:\\Documents\\myfile.pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

            oWordDoc.Close(0);
            oWord.Quit();

            MessageBox.Show("merge successful");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                Outlook.Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);

                string filepath = @"D:\Documents\EMPLOYEE PERFORMANCE REVIEW.docx";

                mail.To = "asanchez@fenceworks.us";
                mail.Subject = "Performance Reviews";
                mail.Body = "benefit enrollments test";
                mail.Importance = Outlook.OlImportance.olImportanceNormal;
                Outlook.Attachment file = mail.Attachments.Add(filepath, Outlook.OlAttachmentType.olByValue, 1, filepath);

                mail.Send();
                MessageBox.Show("message sent");

                _app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
            }
            catch
            {
                MessageBox.Show("failed to send email");
            }
        }
    }
}
