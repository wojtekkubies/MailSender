using MailSenderApp.Class;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace MailSenderApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        JavaScriptSerializer jss = new JavaScriptSerializer();

        DataTable dt = new DataTable();
        BindingSource bd = new BindingSource();
        private void zamknijToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
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
                        // Open the Excel file and load it to the ExcelPackage
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
            string text = "";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                text = File.ReadAllText(file);
            }
            ParseDatabase(text);
        }

        public void ParseDatabase(string text)
        {
            RootObject items = jss.Deserialize<RootObject>(text);
            foreach (var item in items.szablony)
            {
                comboBox1.Items.Add(item.Nazwa);
            }
        }
    }
}
