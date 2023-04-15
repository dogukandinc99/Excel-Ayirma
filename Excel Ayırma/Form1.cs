using System.Diagnostics;

namespace Excel_Ayırma
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        ExcelIslemler excel = new ExcelIslemler();
        String folderpatch = "";

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
            ofd.Multiselect = true;
            ofd.ShowDialog();
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
            foreach (String cell in excel.dict.Keys)
            {
                listBox1.Items.Add(cell.ToString());
            }
        }

        private void listbtn_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = excel.getDataTable();
        }

        private void saveexcelbtn_Click(object sender, EventArgs e)
        {
            //deneme
            Debug.Print("deneme");
            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                excel.excelOpen(ofd.FileNames[i].ToString());
                excel.saveExcel(saveadressfoldertxt.Text, cellvaluetxt.Text + "_" + ofd.SafeFileNames[i].ToString());
                Debug.Print((i + 1) + ". kayıt aktarımı tamamlandı................................................");
            }
            MessageBox.Show("Kayıt işlemi tamamlanmıştır.");
        }

        private void testbtn_Click(object sender, EventArgs e)
        {
            excel.sheetRowSpace();
            dataGridView1.DataSource = excel.getDataTable();
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