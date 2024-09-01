using System.Data;
using System.Diagnostics;
using _Excel = Microsoft.Office.Interop.Excel;

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


        // Gelen adresdeki excel dosyasını açar ve 1. sayfa seçilir.
        public void excelOpen(String path)
        {
            try
            {
                workbook = excel.Workbooks.Open(path);
                worksheet = workbook.Worksheets[1];
                columncontrolnumber = 8;
            }
            catch (Exception e)
            {
                MessageBox.Show("Dosya açılırken bir sorun oluştu. \nHata Kodu: " + e.ToString());
            }

        }


        // Progresbar ilk ayarı için oluşturuldu.
        void progressBarSetting()
        {
            progress.Minimum = 0;
            progress.Maximum = rowsCount - 1;
            progress.Value = 0;
        }


        // Hücreyi satırlara böler ve böldüğü satırlara başlık ekler.
        public void textToColumn()
        {
            Debug.Print("Satırlar sütunlara dönüştürülüyor...");

            _Excel.Range orijinalColumnRange = worksheet.Range["G:G"];
            _Excel.Range newColumnRange = worksheet.UsedRange;
            int rowindex = 1;
            int columnindex = 7;
            try
            {
                foreach (_Excel.Range cell in orijinalColumnRange.Cells)
                {
                    if (cell.Value != null)
                    {
                        String[] cellparts = cell.Value.ToString().Split('\\');

                        foreach (string part in cellparts)
                        {
                            newColumnRange.Cells[rowindex, columnindex].Value2 = part.ToString();
                            columnindex++;
                        }
                        rowindex++;
                        columnindex = 7;
                    }
                    else
                    {
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Hücreler sütunlara bölünürken bir sorun oluştu. \n Hata kodu: " + e.ToString());
            }

            try
            {
                // bölme işlemi yapıldıktan sonra ilk 2 sütun boş geliyordu. boş sürunları sildim.
                for (int i = 0; i < 2; i++)
                {
                    newColumnRange = worksheet.Columns[7];
                    newColumnRange.Delete();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Boş sütunlar silinirkenbir hata oluştu. \n Hata kodu: " + e.ToString());
            }

            try
            {
                // İlk sütun bölme işlemi yaptıktan sonra boş geliyordu. Harfler ile doldurdum.
                int columnnumber = 65;
                rowindex = 1;
                newColumnRange = worksheet.UsedRange;
                for (int i = columnindex; i < (columnindex + 15); i++)
                {
                    String charr = Convert.ToChar(columnnumber).ToString();
                    newColumnRange.Cells[rowindex, i].Value = charr.ToString();
                    columnnumber++;
                }

            }
            catch (Exception e)
            {

                MessageBox.Show("İlk satırda boş olan hücrelere harf eklenirken bir hata oluştu. \n Hata kodu: " + e.ToString());
            }

        }


        // Sürelerde sıfır yazanları 1 e dönüştürür.
        public void zeroChangeOne()
        {
            Debug.Print("Süreleri sıfır olanlar bir ile değiştiriliyor...");

            int rowindex = 2;
            try
            {
                _Excel.Range originalColumn = worksheet.Range["F:F"];
                foreach (_Excel.Range cell in originalColumn.Cells)
                {
                    if (cell.Value != null)
                    {
                        try
                        {
                            if (Convert.ToInt32(cell.Value.ToString()) == 0)
                            {
                                originalColumn.Cells[rowindex, 1].Value = 1;
                            }
                        }
                        catch { continue; }
                    }
                    else break;
                    rowindex += 1;
                }
            }
            catch (Exception e)
            {

                MessageBox.Show("Süresi sıfır olanları bir yaparken sorun oluştu.. \n Hata kodu: " + e.ToString());
            }

            excelquit(true);
        }


        // Adresdeki exceli dataTableList nesnesine aktarır.
        void dataTable()
        {
            try
            {
                //sayfadaki satır ve sütun sayısını değşkenlere aldım.
                rowsCount = worksheet.UsedRange.Rows.Count;
                columnsCount = worksheet.UsedRange.Columns.Count;

                progressBarSetting();

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

                DataView dataView = dataTableList.DefaultView;
                dataView.Sort = dataTableList.Columns[columncontrolnumber].ColumnName + " ASC";
                dataTableList = dataView.ToTable();

                Debug.Print("DataTable nesnesine aktarma başarılı...");
            }
            catch (Exception e)
            {
                MessageBox.Show("Kayıtlar aktarılırken beklenmedik bir hata oluştu. " +
                    "Lütfen teknik birim ile iletişime geçiniz.\n Hata kodu:" + e.Message.ToString());
                excelquit(false);
            }
        }


        // dataTableList nesnesini farklı yerlerde kullanabilmek için oluşturuldu.
        public DataTable getDataTable() { return dataTableList; }


        // Sayfa oluşturmak için aynı değerleri teke indirip diziye ekliyor
        void sheetnamelist()
        {
            //dizi List nesnesinin içini temizler.
            dict.Clear();
            bool start = false;
            int sayi = 0;

            Debug.Print("Sayfa isimleri alınıyor...");
            for (int i = 1; i < dataTableList.Rows.Count; i++)
            {
               string value = dataTableList.Rows[i][columncontrolnumber].ToString();

                if (value != null) // null değer kontrolü
                {                    
                    sayi++;
                    if (!dict.ContainsKey(value))
                    {
                        dict.Add(value, 1);
                        if (start == true)
                        {
                            Debug.Print(dataTableList.Rows[i - 1][columncontrolnumber].ToString() + ": " + sayi);
                        }
                        start = true;
                        sayi = 0;
                    }
                }
                else break;
            }
        }

        // dataTableList nesnesini yeni bir excele kaydeder.
        void newExcel()
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

                //dataTableList nesnesini temizler
                dataTableList.Clear();
                Debug.Print(i.ToString() + ". sayfa DataTable nesnesine aktarımı yapılıyor...");
                dataTable();

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


        // devam İşlem yapılacak exceldeki verileri dataTableList nesnesine aktarır. Yeni excel oluşturur ve yapılması gereken işlmelerden sonra yeni exceli kaydeder.
        public void saveExcel(String adres, String filename)
        {
            dataTableList.Clear();
            dataTable();
            sheetnamelist();
            excelquit(false);
            newExcel();
            sheetRowSpace();
            workbook.SaveAs(@adres + @"\" + filename, _Excel.XlFileFormat.xlWorkbookNormal);
            excelquit(true);
        }


        // Açık olan exceli kapatır.
        public void excelquit(Boolean saveornotsave)
        {
            workbook.Close(saveornotsave);
            excel.Quit();
        }
    }
}