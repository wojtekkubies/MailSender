using Mailer.Class;
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

namespace Mailer.Forms
{
    public partial class ListaSzablonow : Window
    {
        SqliteHelper lhelper;
        Template edytowanyTemplate;

        public ListaSzablonow()
        {
            lhelper = new SqliteHelper();
            InitializeComponent();
            FillTemmplateGrid();
        }

        public void FillTemmplateGrid()
        {
            var r = lhelper.GetAllTemplates();
            TemplateGrid.ItemsSource = lhelper.GetAllTemplates();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (edytowanyTemplate == null)
            {
                MessageBox.Show("Zaznacz rekord do usuniecia.");
            }
            else
            {
                lhelper.Delete(edytowanyTemplate.TemplateId);
                FillTemmplateGrid();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Szablony szablon = new Szablony();
            szablon.Show();
        }

        private void TemplateGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            edytowanyTemplate = TemplateGrid.SelectedItem as Template;
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            FillTemmplateGrid();
        }
    }
}
