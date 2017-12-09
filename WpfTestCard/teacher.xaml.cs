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
using System.Windows.Shapes;
using System.Printing;
using System.Runtime.CompilerServices;

namespace WpfTestCard
{
    /// <summary>
    /// teacher.xaml 的交互逻辑
    /// </summary>
    public partial class teacher : Window
    {
        public teacher()
        {
            InitializeComponent();
        }

        private void print_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PrintDialog printD = new PrintDialog();
                if (printD.ShowDialog() == true)
                {
                    printD.PrintDocument(((IDocumentPaginatorSource)docViewer.Document).DocumentPaginator,"simple"); 
                }
            }
            catch (RuntimeWrappedException dd)
            {
                MessageBox.Show(dd.Message);
                //throw;
            }
            //catch (Exception es)
            //{
            //    MessageBox.Show(es.Message);
            //    //throw;
            //}
            
        }
    }
}
