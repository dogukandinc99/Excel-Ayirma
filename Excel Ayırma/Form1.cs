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

        public void progressbarMax(int max)
        {
            progressBar1.Maximum = max;
        }
        public void progressbarFill()
        {
            progressBar1.Value += 1;
        }
        public void progressbarClear()
        {
            progressBar1.Value = 0;
        }

        private void fileselectbtn_Click(object sender, EventArgs e)
        {
            ofd.Title = "Excel Dosyası Seçiniz.";
            ofd.Filter = "Excel Dosyası |*.xlsx; *.xls";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            ofd.Multiselect = true;
            ofd.ShowDialog();
            adresstxt.Text = ofd.FileName;
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

        private void saveexcelbtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                excel.excelOpen(ofd.FileNames[i].ToString());
                dataGridView1.DataSource = excel.getDataTable();
                excel.saveExcel(saveadressfoldertxt.Text, cellvaluetxt.Text + "_" + ofd.SafeFileNames[i].ToString());
                Debug.Print((i + 1) + " kayıdın aktarımı tamamlandı................................................");
            }
            MessageBox.Show("Kayıt işlemi tamamlanmıştır.");
        }
    }
}