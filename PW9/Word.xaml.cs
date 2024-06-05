using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace PW9
{
    /// <summary>
    /// Логика взаимодействия для Word.xaml
    /// </summary>
    public partial class Word : Window
    {
        public Word()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            var fs = new FileStream(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Конвертировали.rtf", FileMode.Create);
            range.Save(fs, DataFormats.Rtf);
            fs.Close();

            Document doc = new Document();
            doc.LoadFromFile(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Конвертировали.rtf");
            doc.SaveToFile(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Из проги.docx", FileFormat.Docx);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Document doc = new Document();
            doc.LoadFromFile(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Документ Microsoft Word.docx");
            doc.SaveToFile(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Конвертировали.rtf", FileFormat.Rtf);

            var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            var fs = new FileStream(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Конвертировали.rtf", FileMode.OpenOrCreate);
            range.Load(fs, DataFormats.Rtf);
            fs.Close();
        }
    }
}
