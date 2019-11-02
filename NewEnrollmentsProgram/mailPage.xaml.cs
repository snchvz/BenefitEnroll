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
using System.Runtime.InteropServices;

namespace NewEnrollmentsProgram
{
    /// <summary>
    /// Interaction logic for mailPage.xaml
    /// </summary>
    public partial class mailPage : Page
    {
        public mailPage()
        {
            InitializeComponent();
        }

        private void MailButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Outlook.Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);

                string filepath = @"D:\Documents\EMPLOYEE PERFORMANCE REVIEW.docx";

                mail.To = "mchavez@fenceworks.us";
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
