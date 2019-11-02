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
    /// Interaction logic for MergePage.xaml
    /// </summary>
    public partial class MergePage : Page
    {
        public MergePage()
        {
            InitializeComponent();

            CompanyComboBox.Items.Add("FWI");
            CompanyComboBox.Items.Add("FSI");
            CompanyComboBox.Items.Add("FCI");
            CompanyComboBox.Items.Add("ACFS");


            //TODO-- Set up as an observable collection to use raiseProperty events for when the company is changed dynamically
            YardComboBox.Items.Add("All Yards");
            YardComboBox.Items.Add("AN");
            YardComboBox.Items.Add("BK");
            YardComboBox.Items.Add("FR");
            YardComboBox.Items.Add("MN");
            YardComboBox.Items.Add("OC");
            YardComboBox.Items.Add("PD");
            YardComboBox.Items.Add("RV");
            YardComboBox.Items.Add("SA");
            YardComboBox.Items.Add("SJ");
            YardComboBox.Items.Add("SP");


            DocMergeComboBox.Items.Add("Performance Review");
            DocMergeComboBox.Items.Add("Payroll Deductions");
            DocMergeComboBox.Items.Add("New Enrollment Memos");

            DetectCompany();
        }

        public void DetectCompany()
        {
            switch(CompanyStatic.Instance.companyName)
            {
                case "FWI":
                    CompanyComboBox.SelectedIndex = 0;
                    break;
                case "FSI":
                    CompanyComboBox.SelectedIndex = 1;
                    break;
                case "FCI":
                    CompanyComboBox.SelectedIndex = 2;
                    break;
                case "ACFS":
                    CompanyComboBox.SelectedIndex = 3;
                    break;
                default:
                    break;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (DocMergeComboBox.SelectedItem == null || CompanyComboBox.SelectedItem == null)
            {
                MessageBox.Show("Please select a company, yard and document type to merge");
                MessageBox.Show(CompanyStatic.Instance.companyName);
                return;
            }

            Object oTemplatePath;

            //TODO-- clean up switch case and seperate into another function
            switch (CompanyComboBox.SelectedIndex)
            {
                case 0:
                    switch (DocMergeComboBox.SelectedIndex)
                    {
                        case 0:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\FWI Performance Review Template.docx";
                            break;
                        case 1:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Payroll Deduction Templates\FWI Payroll Deduction Template.docx";
                            break;
                        case 2:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Memos\FWI New Enrollment Memo.doc";
                            break;
                        default:
                            return;
                    }
                    break;
                case 1:
                    switch (DocMergeComboBox.SelectedIndex)
                    {
                        case 0:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\FSI Performance Review Template.docx";
                            break;
                        case 1:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Payroll Deduction Templates\FSI Payroll Deduction Template.docx";
                            break;
                        case 2:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Memos\FSI New Enrollment Memo.doc";
                            break;
                        default:
                            return;
                    }
                    break;
                case 2:
                    switch (DocMergeComboBox.SelectedIndex)
                    {
                        case 0:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\FCI Performance Review Template.docx";
                            break;
                        case 1:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Payroll Deduction Templates\FCI Payroll Deduction Template.docx";
                            break;
                        case 2:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Memos\FCI New Enrollment Memo.doc";
                            break;
                        default:
                            return;
                    }
                    break;
                case 3:
                    switch (DocMergeComboBox.SelectedIndex)
                    {
                        case 0:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\ACFS Performance Review Template.docx";
                            break;
                        case 1:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Payroll Deduction Templates\ACFS Payroll Deduction Template.docx";
                            break;
                        case 2:
                            oTemplatePath = @"D:\Documents\New Enroll Memos\Memos\ACFS New Enrollment Memo.doc";
                            break;
                        default:
                            return;
                    }
                    break;
                default:
                    return;
            }

            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            try
            {
                oWordDoc = oWord.Documents.Open(oTemplatePath);
            }

            catch
            {
                MessageBox.Show("could not find " + oTemplatePath.ToString());
                //oWordDoc.Close(0);
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

            object qry = "select * from [Sheet1$]";
            Object oMissing = System.Reflection.Missing.Value;

            if (YardComboBox.SelectedValue != null )
            {
                if(YardComboBox.SelectedValue.ToString() != "All Yards")
                    qry += " WHERE Yard = '" + YardComboBox.SelectedValue.ToString() + "'";
            }                

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

            try
            {
                oWordDoc.MailMerge.Execute();
            }
            catch
            {
                oWordDoc.Close(0);
                oWord.Quit();

                Marshal.ReleaseComObject(oWord);
                Marshal.ReleaseComObject(oWordDoc);

                MessageBox.Show("The datasource contains no records");

                return;
            }  
            
            oWordDoc.Close(0);
            oWord.Quit();

            Marshal.ReleaseComObject(oWord);
            Marshal.ReleaseComObject(oWordDoc);

            MessageBox.Show("merge successful");
        }        
    }
}
