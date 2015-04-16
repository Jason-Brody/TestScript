using SAPTestScripts;
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
using winForm = System.Windows.Forms;

namespace GLReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            winForm.OpenFileDialog ofd = new winForm.OpenFileDialog();
            if(ofd.ShowDialog()==winForm.DialogResult.OK)
            {
                tb_File.Text = ofd.FileName;
            }
            
        }

        private void btn_Run_Click(object sender, RoutedEventArgs e)
        {
            winForm.SaveFileDialog sfd = new winForm.SaveFileDialog();
            if(sfd.ShowDialog() == winForm.DialogResult.OK)
            {
                RevaluationOfGLAccount script = new RevaluationOfGLAccount();
                script.Read(tb_File.Text);
                script.GetReport(tb_Date.Text,sfd.FileName);
            }
        }
    }
}
