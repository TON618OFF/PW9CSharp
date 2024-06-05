using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PW9
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

        private void create_word_Click(object sender, RoutedEventArgs e)
        {
            Word word = new Word();
            word.Show();
            Close();
        }

        private void create_excel_Click(object sender, RoutedEventArgs e)
        {
            Excel excel = new Excel();
            excel.Show();
            Close();
        }

    }
}