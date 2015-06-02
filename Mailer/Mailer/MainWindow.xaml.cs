using Mailer.Class;
using Mailer.Forms;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Mailer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SqliteHelper lhelper;
        DataTable dt = new DataTable();
        MailHelper mailHelper;
        ConfigHelper configHelper;
        public string szablonDoWyslania = "";

        public MainWindow()
        {
            InitializeComponent();
            lhelper = new SqliteHelper();
            mailHelper = new MailHelper();
            configHelper = new ConfigHelper();
            lhelper.CreateDbAndTableIfNotExist();
            AktualizujExcelLabel();
            SprobujZaladowacPlikZDomyslnejLokalizacji();
        }

        public void SprobujZaladowacPlikZDomyslnejLokalizacji()
        {
            if (configHelper.AktualnaSciezkaExcela.Length != 0)
            {
                try
                {
                    ZaladujExcela(configHelper.AktualnaSciezkaExcela);
                }
                catch
                {

                }
            }
        }

        public void AktualizujExcelLabel()
        {
            if (configHelper.AktualnaSciezkaExcela.Length == 0)
            {
                PathLabel.Content = "Brak domyślnej ścieżki.";
            }
            else
            {
                PathLabel.Content = "Sciezka excela: " + configHelper.AktualnaSciezkaExcela;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ListaSzablonow lSzablonow = new ListaSzablonow();
            lSzablonow.Show();
        }
       
        public void ZaladujExcela(string path)
        {
            try
            {
                using (ExcelPackage pck = new ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                    dt = WorksheetToDataTable(ws, true);
                    ExcelGrid.ItemsSource = dt.DefaultView;                    
                }
                configHelper.PodmienUstawieniaSciezkiExcela(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed. Original error: " + ex.Message);
            }
        }

        private DataTable WorksheetToDataTable(ExcelWorksheet ws, bool hasHeader = true)
        {
            dt = new DataTable(ws.Name);
            int totalCols = ws.Dimension.End.Column;
            int totalRows = ws.Dimension.End.Row;
            int startRow = hasHeader ? 2 : 1;
            ExcelRange wsRow;
            DataRow dr;
            dt.Columns.Add(new DataColumn("Wyślij", typeof(bool)));
            foreach (var firstRowCell in ws.Cells[1, 1, 1, totalCols])
            {
                dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
            {
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                dr[0] = true;
                foreach (var cell in wsRow)
                {                    
                    dr[cell.Start.Column] = cell.Text;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xlsx";
            openfile.Filter = "(.xlsx)|*.xlsx";
            var browsefile = openfile.ShowDialog();
            if (browsefile == true)
            {
                ZaladujExcela(openfile.FileName);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            WczytajDaneDoTemplateComboBox();
            
        }

        public void WczytajDaneDoTemplateComboBox()
        {
            SzablonyComboBox.Items.Clear();
            List<Template> res = new List<Template>();
            res = lhelper.GetAllTemplates();
            foreach (var item in res)
            {
                SzablonyComboBox.Items.Add(item.TemplateName);
            }
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            WczytajDaneDoTemplateComboBox();
        }

        private void SzablonyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<Template> res = new List<Template>();
            res = lhelper.GetAllTemplates();
            foreach (var item in res)
            {
                if (item.TemplateName == SzablonyComboBox.SelectedItem.ToString())
                {
                    szablonDoWyslania = item.TemplateText;
                }
            }
        }

        public void wyslijEmaile()
        {
            Dictionary<string, string> slownikWartosci = new Dictionary<string, string>();
            var row = GetDataGridRows(ExcelGrid);
            foreach (DataGridRow r in row)
            {
                DataRowView rv = (DataRowView)r.Item;                
                bool t = bool.Parse(rv.Row[0].ToString());
                if (t == true)
                {
                    for (int i = 1; i < ExcelGrid.Columns.Count; i++)
                    {
                        string kolumna = ExcelGrid.Columns[i].Header.ToString();
                        string wartosc = rv.Row[i].ToString(); 
                        slownikWartosci[kolumna] = wartosc;
                    }
                    mailHelper.sparsujIWyslij(slownikWartosci, szablonDoWyslania);
                }
            }
        }

        public IEnumerable<DataGridRow> GetDataGridRows(DataGrid grid)
        {
            var itemsSource = grid.ItemsSource as IEnumerable;
            if (null == itemsSource) yield return null;
            foreach (var item in itemsSource)
            {
                var row = grid.ItemContainerGenerator.ContainerFromItem(item) as DataGridRow;
                if (null != row) yield return row;
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            wyslijEmaile();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

        }
    }
}
