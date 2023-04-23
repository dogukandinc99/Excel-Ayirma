using System.Data;
using System.Diagnostics;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Excel_Ayırma
{
    public class ExcelIslemler
    {
        _Excel.Application excel = new _Excel.Application();
        _Excel.Workbook workbook;
        _Excel.Worksheet worksheet;

        System.Data.DataTable dataTableList = new System.Data.DataTable("Excel-List");

        public Dictionary<string, int> dict = new Dictionary<string, int>();
        int columncontrolnumber = 8;
        int rowsCount = 0, columnsCount = 0;
        ProgressBar progress;
        Label label1;

        public ExcelIslemler(ProgressBar progress, Label label)
        {
            this.progress = progress;
            this.label1 = label;
        }

        // Gelen adresdeki excel dosyasını açar ve dataTable methodunu çalıştırır.
        // dataTable methodu çalıştıkdan sonra adresdeki exceli kapatır.
        public void excelOpen(String path)
        {
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[1];

            //dizi List nesnesinin içini temizler.
            dict.Clear();

            //kontrol edilecek sütun
            columncontrolnumber = 8;
            defaultValue();
            sheetnamelist();
            excelquit();
        }


        // Exceli açtıktan sonra başka işlemler için yeniden çağırmam gerektiğinden dolayı excelOpen methodunu oluşturduö.
        void defaultValue()
        {
            //sayfadaki satır ve sütun sayısını değşkenlere aldım.
            rowsCount = worksheet.UsedRange.Rows.Count;
            columnsCount = worksheet.UsedRange.Columns.Count;

            progress.Minimum = 0;
            progress.Maximum = rowsCount - 1;
            progress.Value = 0;

            //dataTableList nesnesini temizler
            dataTableList.Clear();
            dataTable();
        }


        // Adresdeki exceli dataTableList nesnesine aktarır.
        void dataTable()
        {
            try
            {
                // dataTableList nesnesine column oluşturuyor.
                if (dataTableList.Columns.Count != columnsCount)
                {
                    for (int i = 0; i < columnsCount; i++)
                    {
                        dataTableList.Columns.Add(worksheet.Cells[1, i + 1].Value2.ToString());
                    }
                }

                // dataTableList nesnesine exceldeki değerleri satır satır ekliyor.
                for (int i = 2; i < rowsCount + 1; i++)
                {
                    DataRow dataRow = dataTableList.NewRow();
                    for (int j = 1; j < columnsCount; j++)
                    {
                        dataRow[j - 1] = worksheet.Cells[i, j].Value;
                    }
                    dataTableList.Rows.Add(dataRow);
                    progress.Value += 1;
                    label1.Text = progress.Value.ToString() + " /" + progress.Maximum.ToString();
                }
                Debug.Print("DataTable nesnesine aktarma başarılı...");
            }
            catch (Exception e)
            {
                MessageBox.Show("Kayıtlar aktarılırken beklenmedik bir hata oluştu. " +
                    "Lütfen teknik birim ile iletişime geçiniz.\n Hata kodu:" + e.Message.ToString());
                excelquit();
            }
        }


        // dataTableList nesnesini farklı yerlerde kullanabilmek için oluşturuldu.
        public DataTable getDataTable() { return dataTableList; }


        // Sayfa oluşturmak için aynı değerleri teke indirip diziye ekliyor
        void sheetnamelist()
        {
            _Excel.Range range;

            Debug.Print("Sayfa isimleri alınıyor...");
            for (int i = 2; i < worksheet.Rows.Count + 1; i++)
            {
                range = worksheet.Cells[i, columncontrolnumber + 1];
                if (range.Value != null) // null değer kontrolü
                {
                    string value = range.Value;
                    Debug.Print(value.ToString());
                    if (!dict.ContainsKey(value))
                    {
                        dict.Add(value, 1);
                    }
                }
                else break;
            }
        }


        // dataTableList nesnesini yeni bir excele kaydeder.
        void newExcel(String adres, String filename)
        {
            workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            foreach (String item in dict.Keys)
            {
                // Yeni açılan excelde sayfalar oluşturur.
                worksheet = workbook.Worksheets.Add();
                worksheet.Name = sheetnamelenght(item.ToString());
                worksheet = workbook.Worksheets[1];

                Debug.Print("Sayfalar oluşturuluyor...");
                for (int j = 0; j < dataTableList.Columns.Count; j++)
                {
                    worksheet.Cells[1, j + 1] = dataTableList.Columns[j].ColumnName.ToString();
                }

                // Açılan sayfanın ismine göre sayfanın içine satırları koyar
                int row = 1;
                for (int i = 0; i < dataTableList.Rows.Count; i++)
                {
                    if (dataTableList.Rows[i][columncontrolnumber].ToString() == item.ToString())
                    {
                        for (int j = 0; j < dataTableList.Columns.Count; j++)
                        {
                            worksheet.Cells[row + 1, j + 1] = dataTableList.Rows[i][j];
                        }
                        row++;
                    }
                }
            }
            worksheet = workbook.Worksheets[workbook.Worksheets.Count];
            worksheet.Delete();
        }


        // Sayfa adı uzun ise ilk 15 karakteri alıyor.
        String sheetnamelenght(String value)
        {
            String control;
            if (value.Length < 32)
            {
                control = value;
            }
            else
            {
                control = value.Substring(0, 31).ToString();
            }
            return control;
        }


        // Kaydedilecek yeni excelde sayfa sayfa gezerek sütuna göre farklı olan satırların arasına boşluk bırakıp yeni excele aktarır.
        public void sheetRowSpace()
        {
            String sheetname = "";
            for (int i = 1; i <= workbook.Worksheets.Count; i++)
            {
                worksheet = workbook.Worksheets[i];
                defaultValue();

                sheetname = worksheet.Name;
                switch (sheetname)
                {
                    case "A HABER":
                        columncontrolnumber = 10;
                        break;
                    case "A SPOR":
                        columncontrolnumber = 10;
                        break;
                    case "APARA":
                        columncontrolnumber = 10;
                        break;
                    case "ATV":
                        columncontrolnumber = 9;
                        break;
                    case "VAV":
                        columncontrolnumber = 9;
                        break;
                    case "TEKNIK BILGI ISLEM":
                        columncontrolnumber = 11;
                        break;
                    case "GENEL ARŞİV (AJANSLAR -İNGESTLE":
                        columncontrolnumber = 12;
                        break;
                    default:
                        columncontrolnumber = 8;
                        break;
                }

                Debug.Print("Sıralama yapılıyor...");
                DataView dataView = dataTableList.DefaultView;
                dataView.Sort = dataTableList.Columns[columncontrolnumber].ColumnName + " ASC";
                dataTableList = dataView.ToTable();

                Debug.Print("Farklı olan satırlar ayrılıyor...");
                string prevValue = null;
                for (int j = 0; j < dataTableList.Rows.Count; j++)
                {
                    string currentValue = dataTableList.Rows[j][columncontrolnumber].ToString();

                    if (string.IsNullOrEmpty(currentValue))
                    {
                        continue;
                    }

                    if (prevValue == currentValue)
                    {
                        continue;
                    }

                    dataTableList.Rows.InsertAt(emptyRowSpace(), j);
                    prevValue = currentValue;
                }

                for (int j = 0; j < dataTableList.Rows.Count; j++)
                {
                    for (int k = 0; k < dataTableList.Columns.Count; k++)
                    {
                        worksheet.Cells[j + 2, k + 1] = dataTableList.Rows[j][k].ToString();
                    }
                }
            }
        }


        // Boş satır oluşturur.
        DataRow emptyRowSpace()
        {
            DataRow dr = dataTableList.NewRow();
            for (int j = 0; j < columnsCount; j++)
            {
                dr[j] = "";
            }
            return dr;
        }


        // Yeni exceli kaydeder.
        public void saveExcel(String adres, String filename)
        {
            newExcel(adres, filename);
            sheetRowSpace();
            workbook.SaveAs(@adres + @"\" + filename, _Excel.XlFileFormat.xlWorkbookNormal);
            excelquit();
        }


        // Açık olan exceli kapatır.
        public void excelquit()
        {
            workbook.Close();
            excel.Quit();
        }
    }
}
