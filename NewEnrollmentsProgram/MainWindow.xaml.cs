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
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        MergePage mergePage = new MergePage();
        MakeListPage listPage = new MakeListPage();
        mailPage mail_Page = new mailPage();
        
        

        public MainWindow()
        {
            InitializeComponent();
            CompanyStatic.Instance.companyName = "";
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Main.Content = mergePage;
            mergePage.DetectCompany();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Main.Content = listPage;
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Main.Content = mail_Page;
        }
    }
}
