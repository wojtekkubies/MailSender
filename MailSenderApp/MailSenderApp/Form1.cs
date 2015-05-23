using MailSenderApp.Class;
using MailSenderApp.Forms;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace MailSenderApp
{
    public partial class Form1 : Form
    {
        LiteHelper lhelper;
        public Form1()
        {
            InitializeComponent();
            lhelper = new LiteHelper();
            lhelper.CreateDbAndTableIfNotExist();
        }

        public bool IsValidEmail(string emailaddress)
        {
            try
            {
                MailAddress m = new MailAddress(emailaddress);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        public DataTable themesDataTable = new DataTable();
        DataTable dt = new DataTable();
        BindingSource bd = new BindingSource();
        public string szablonDoWyslania = "";

        private void zamknijToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
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
                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed. Original error: " + ex.Message);
            }
        }

        private static void PodmienUstawieniaSciezkiExcela(string value)
        {
            string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configFile = System.IO.Path.Combine(appPath, "App.config");
            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = configFile;
            System.Configuration.Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
            config.AppSettings.Settings["ExcelPlik"].Value = value;
            config.Save();
        }

        private void wczytajDaneZXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel File (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (ExcelPackage pck = new ExcelPackage())
                    {
                        using (var stream = File.OpenRead(openFileDialog1.FileName))
                        {
                            pck.Load(stream);
                        }

                        ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                        dt = WorksheetToDataTable(ws, true);
                        dataGridView1.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Import failed. Original error: " + ex.Message);
                }
                PodmienUstawieniaSciezkiExcela(openFileDialog1.FileName);
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
            foreach (var firstRowCell in ws.Cells[1, 1, 1, totalCols])
            {
                dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
            {
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    dr[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(dr);
            }
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                for (int i = 1; i < dataGridView1.Rows[j].Cells.Count; i++)
                {
                    if (dataGridView1.Rows[j].Cells.Count == 1) continue;
                    MessageBox.Show(dataGridView1.Rows[j].Cells[i].Value.ToString());
                }
            }
        }

        private void załadujPlikSzablonówToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        public void ParseDatabase()
        {
            comboBox1.Items.Clear();
            List<Template> res = new List<Template>();
            res = lhelper.GetAllTemplates();
            foreach (var item in res)
            {
                comboBox1.Items.Add(item.TemplateName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            wyslijEmaile();
        }

        public void sparsujIWyslij(Dictionary<string, string> slownikWartosci, string tekst)
        {
            char[] separators = new char[] { ' ','\t','\n' };
            string[] tab = tekst.Split(separators);
            string email = "";
            StringBuilder tekstDoWyslania = new StringBuilder();
            foreach (var item in tab)
            {
                if (item.Length > 0)
                {
                    if (item[0] == '$')
                    {
                        foreach (var item2 in slownikWartosci)
                        {
                            if (item2.Key == item.Substring(1, item.Length - 1))
                            {
                                tekstDoWyslania.Append(" " + item2.Value + " ");
                            }
                        }
                    }
                    else
                    {
                        tekstDoWyslania.Append(" " + item + " ");
                    }
                }               
            }

            foreach (var item in slownikWartosci)
            {
                if (IsValidEmail(item.Value))
                {
                    email += item.Value + ";";
                }
            }
            email = email.Substring(0, email.Length - 2);
            string command = "mailto:" + email + "?subject=Test&body=" + tekstDoWyslania.ToString();
            Process.Start(command);
        }

        public void wyslijEmaile()
        {
            Dictionary<string, string> slownikWartosci = new Dictionary<string, string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[SendColumn.Name].Value) == true)
                {
                    for (int i = 1; i < row.Cells.Count; i++)
                    {
                        slownikWartosci[dataGridView1.Columns[i].HeaderText] = row.Cells[i].Value.ToString();
                    }
                    sparsujIWyslij(slownikWartosci, szablonDoWyslania);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<Template> res = new List<Template>();
            res = lhelper.GetAllTemplates();
            foreach (var item in res)
            {
                if (item.TemplateName == comboBox1.SelectedItem.ToString())
                {
                    szablonDoWyslania = item.TemplateText;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sciezka = ConfigurationManager.AppSettings["ExcelPlik"];
            if (sciezka.Trim().Length > 0)
            {
                ZaladujExcela(sciezka.Trim());
            }
            ParseDatabase();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ListaSzablonow szablony = new ListaSzablonow();
            szablony.Show();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            ParseDatabase();
        }
    }
}
