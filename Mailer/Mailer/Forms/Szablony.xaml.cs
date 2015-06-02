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
    /// <summary>
    /// Interaction logic for Szablony.xaml
    /// </summary>
    public partial class Szablony : Window
    {
        SqliteHelper lhelper = new SqliteHelper();
        
        public Szablony()
        {            
            InitializeComponent();
            ErrorLabel.Content = "";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string nazwa = NazwaTxtBox.Text.Trim();
                string text = Text.Text.Trim();
                if (nazwa.Trim().Length == 0 ||
                    text.Trim().Length == 0)
                {
                    ErrorLabel.Content = "Podaj odpowiednia dlugosc tekstów.";
                }
                else
                {
                    lhelper.AddNewTemplate(nazwa, text);
                    NazwaTxtBox.Text = "";
                    Text.Text = "";
                    ErrorLabel.Content = "Dane zostały zapisane.";
                }
            }
            catch
            {
                ErrorLabel.Content = "Nie udalo sie zapisac danych, skontaktuj sie z twórcą.";
            }
        }
    }
}
