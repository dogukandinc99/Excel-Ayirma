using System.Diagnostics;

namespace Excel_Ayırma
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        ExcelIslemler excel;
        String folderpatch = "";

        public Form1()
        {
            InitializeComponent();
            excel = new ExcelIslemler(progressBar1, label4);
        }
        private void fileselectbtn_Click(object sender, EventArgs e)
        {
            ofd.Title = "Excel Dosyası Seçiniz.";
            ofd.Filter = "Excel Dosyası |*.xlsx; *.xls";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            ofd.Multiselect = true;
            ofd.ShowDialog();

            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                adresstxt.Text += ofd.FileNames[i].ToString() + Environment.NewLine;
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

        private void saveexcelbtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                excel.excelOpen(ofd.FileNames[i].ToString());
                dataGridView1.DataSource = excel.getDataTable();
                Debug.Print(ofd.SafeFileNames[i].ToString() + " adlı dosya için işlem başlatılıyor...");
                excel.saveExcel(saveadressfoldertxt.Text, cellvaluetxt.Text + "_" + ofd.SafeFileNames[i].ToString());
                Debug.Print((i + 1) + " kayıdın aktarımı tamamlandı................................................");
            }
            MessageBox.Show("Kayıt işlemi tamamlanmıştır.");
        }

        private void rdybtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ofd.FileNames.Length; i++)
            {
                excel.excelOpen(ofd.FileNames[i].ToString());
                excel.textToColumn();
                excel.zeroChangeOne();
                Debug.Print((i + 1) + " adlı dosya hazırlandı................................................");
            }
            MessageBox.Show("Excel dosyaları ayırma işlemi için hazırlandı.");
        }
    }
}