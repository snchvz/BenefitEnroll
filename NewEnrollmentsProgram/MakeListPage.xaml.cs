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
using System.Windows.Forms;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace NewEnrollmentsProgram
{
    /// <summary>
    /// Interaction logic for MakeListPage.xaml
    /// </summary>
    public partial class MakeListPage : Page
    {
        OpenFileDialog ofd = new OpenFileDialog();
        OpenFileDialog ofdDest = new OpenFileDialog();

        public MakeListPage()
        {
            InitializeComponent();
            MonthComboBox.Items.Add("1");
            MonthComboBox.Items.Add("2");
            MonthComboBox.Items.Add("3");
            MonthComboBox.Items.Add("4");
            MonthComboBox.Items.Add("5");
            MonthComboBox.Items.Add("6");
            MonthComboBox.Items.Add("7");
            MonthComboBox.Items.Add("8");
            MonthComboBox.Items.Add("9");
            MonthComboBox.Items.Add("10");
            MonthComboBox.Items.Add("11");
            MonthComboBox.Items.Add("12");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "xls|*.xlsm;*.xlsx";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                //textBox2.Text = ofd.SafeFileName;
            }
        }

        private void Read_File_Click(object sender, RoutedEventArgs e)
        {
            if (textBox1.Text.Length < 3 || textBox2.Text.Length < 3)
            {
                System.Windows.MessageBox.Show("Please Select a roster excel File and Destination File");
                return;
            }
            else if (MonthComboBox.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Please select month of hire");
                return;
            }

            else if (YearTextBox.Text.Length != 4)
            {
                System.Windows.MessageBox.Show("Please enter a valid year (YYYY)");
                return;
            }

            int month = int.Parse(MonthComboBox.SelectedItem.ToString());


            ExcelRead XL = new ExcelRead(ofd.FileName, 1);
            System.Windows.MessageBox.Show(XL.readWriteCell(month, YearTextBox.Text, ofd.FileName, 1, ofd.SafeFileName, textBox2.Text).ToString());
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ofdDest.Filter = "xls|*.xlsm;*.xlsx";
            if (ofdDest.ShowDialog() == DialogResult.OK)
            {
                //textBox1.Text = ofd.FileName;
                textBox2.Text = ofdDest.FileName;
            }
        }
    }
}
