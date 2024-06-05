using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
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
    /// Логика взаимодействия для Excel.xaml
    /// </summary>
    public partial class Excel : Window
    {
        public Excel()
        {
            InitializeComponent();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            var DataTable = dg.ItemsSource as DataView;

            Workbook wb = new Workbook(); // создание книги
            wb.Worksheets.Clear();
            Worksheet sheesh = wb.Worksheets.Add("Лист 1"); // добавление листа (страницы) в книгу
            sheesh.InsertDataView(DataTable, true, 1, 1); // указываем таблицу (DataTable), есть ли названия колонок, индекс строки, индекс колонки.
            wb.SaveToFile(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Лист Microsoft Excel.xlsx", FileFormat.Version2016);
        }

        private void Load_Click(object sender, RoutedEventArgs e)
        {
            Workbook wb = new Workbook(); // то же самое, что и с word, только Workbook
            wb.LoadFromFile(@"D:\Vyacheslav\Study\Project\Visual C#\WPFs\PW9\контент для проги\Лист Microsoft Excel.xlsx"); // тупо считываем файл

            // с помощью этого мы сможем сделать DataTable - таблицу с данными.
            Worksheet sheesh = wb.Worksheets[0]; // берём первую страницу
            CellRange locatedRange = sheesh.AllocatedRange; // берём область ячеек, в которых есть данные

            // Берём таблицу с помощью экспорта. Помещаем область ячеек и говорим, что первая строка - имена столбцов
            var DataTable = sheesh.ExportDataTable(locatedRange, true);
            dg.ItemsSource = DataTable.DefaultView;
        }

        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            // Создать новый столбец DataGrid
            DataGridTextColumn newColumn = new DataGridTextColumn();

            // Настроить свойства нового столбца
            newColumn.Header = "Новая Колонка";
            newColumn.Binding = new Binding("Новое Свойство");

            // Добавить новый столбец в коллекцию Columns DataGrid
            dg.Columns.Add(newColumn);
        }
    }
}
