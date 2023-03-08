using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using System.Data;
using System.Text.Json.Serialization;
using System.Windows.Forms;


namespace Excel_Ayırma
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        ExcelIslemler excel = new ExcelIslemler();
        String filepath = "", folderpatch = "";

        public Form1()
        {
            InitializeComponent();
        }

        void progressbarfill(ProgressBar progressBar, int min, int max)
        {
            progressBar.Minimum = min;
            progressBar.Maximum = max;
        }
        private void fileselectbtn_Click(object sender, EventArgs e)
        {

            ofd.Title = "Excel Dosyası Seçiniz.";
            ofd.Filter = "Excel Dosyası |*.xlsx; *.xls";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filepath = ofd.FileName;
                adresstxt.Text = filepath;
                excel.excelOpen(filepath);
                cellvaluebtn.Enabled = true;
                cellvaluetxt.Enabled = true;
                saveselectedfolderbtn.Enabled = true;
                listbtn.Enabled = true;
                excelclosebtn.Enabled = true;
            }
        }

        private void saveselectedfolderbtn_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                saveexcelbtn.Enabled = true;
                folderpatch = fbd.SelectedPath;
                saveadressfoldertxt.Text = folderpatch;
            }
        }

        private void cellvaluebtn_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            for (int i = 0; i < excel.dizi.Length; i++)
            {
                if (excel.dizi[i] != null)
                {
                    listBox1.Items.Add(excel.dizi[i].ToString());
                }
                else
                {
                    break;
                }
            }
        }

        private void listbtn_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = excel.getDataTable();
        }

        private void saveexcelbtn_Click(object sender, EventArgs e)
        {
            excel.saveExcel(saveadressfoldertxt.Text, cellvaluetxt.Text + "_" + ofd.SafeFileName.ToString());
        }

        private void excelclosebtn_Click(object sender, EventArgs e)
        {
            excel.excelquit();
            fileselectbtn.Enabled = true;
            cellvaluebtn.Enabled = false;
            cellvaluetxt.Enabled = false;
            listbtn.Enabled = false;
            saveselectedfolderbtn.Enabled = false;
            saveexcelbtn.Enabled = false;
        }
    }
}