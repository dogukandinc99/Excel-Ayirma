using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Ayırma
{
    public class ExcelIslemler
    {
        _Excel.Application excel = new _Excel.Application();
        _Excel.Workbook workbook;
        _Excel.Worksheet worksheet;

        System.Data.DataTable dataTableList = new System.Data.DataTable("Excel-List");

        public String[] dizi = new String[10];
        int columncontrolnumber = 8;
        int rowsCount = 0, columnsCount = 0;

        // Gelen adresdeki excel dosyasını açar ve dataTable methodunu çalıştırır. dataTable methodu çalıştıkdan sonra adresdeki exceli kapatır.
        public void excelOpen(String path)
        {
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[1];
            defaultValue();
            sheetnamelist();
            excelquit();
        }


        // Exceli açtıktan sonra başka işlemler için yeniden çağırmam gerektiğinden dolayı excelOpen methodundan ayırdım.
        void defaultValue()
        {
            rowsCount = worksheet.UsedRange.Rows.Count;
            columnsCount = worksheet.UsedRange.Columns.Count;
            dataTableList.Clear();
            dataTable();
        }


        // Exceldeki hücreyi getirir. Eğer hücre boş ise geriye "Empty" değerini dönderir.
        String getReadCell(int i, int j)
        {
            string value;
            i++;
            j++;
            if (worksheet.Cells[i, j].Value2 != null)
            {
                value = worksheet.Cells[i, j].Value2.ToString();
            }
            else
            {
                value = "Empty";
            }
            return value;
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
                        dataTableList.Columns.Add(getReadCell(0, i));
                    }
                }
                for (int i = 2; i < rowsCount + 1; i++)
                {
                    DataRow dataRow = dataTableList.NewRow();
                    for (int j = 1; j < columnsCount; j++)
                    {
                        dataRow[j - 1] = worksheet.Cells[i, j].Value;
                    }
                    dataTableList.Rows.Add(dataRow);
                }
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
            String cellvalue = "";
            int sayac = 0, control = 0;
            try
            {
                for (int i = 1; i < rowsCount - 1; i++)
                {
                    cellvalue = getReadCell(i, columncontrolnumber);

                    if (cellvalue == getReadCell(i - 1, columncontrolnumber) || getReadCell(i - 1, columncontrolnumber) == "")
                    {

                    }
                    else
                    {
                        String sheetname = cellvalue;
                        for (int j = 0; j < dizi.Length; j++)
                        {
                            if (dizi[j] == sheetname)
                            {
                                control++;
                            }
                        }
                        if (control == 0)
                        {
                            dizi[sayac] = sheetname;
                            sayac += 1;
                        }
                        control = 0;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("hata mesajı:" + e.ToString());
            }
        }


        // dataTableList nesnesini yeni bir excele kaydeder.
        void newExcel(String adres, String filename)
        {
            workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            for (int i = 0; i < dizi.Length; i++)
            {
                if (dizi[i] != null)
                {
                    worksheet = workbook.Worksheets.Add();
                    worksheet.Name = sheetnamelenght(dizi[i].ToString());
                    worksheet = workbook.Worksheets[1];
                    for (int j = 0; j < dataTableList.Columns.Count; j++)
                    {
                        worksheet.Cells[1, j + 1] = dataTableList.Columns[j].ColumnName.ToString();
                    }
                }
                else
                    break;
            }

            int row;
            for (int i = 0; i < dizi.Length; i++)
            {
                if (dizi[i] != null)
                {
                    worksheet = workbook.Worksheets[sheetnamelenght(dizi[i]).ToString()];
                    row = 1;
                    for (int j = 0; j < dataTableList.Rows.Count; j++)
                    {
                        if (dataTableList.Rows[j][columncontrolnumber].ToString() == dizi[i])
                        {
                            for (int k = 0; k < dataTableList.Columns.Count; k++)
                            {
                                worksheet.Cells[row + 1, k + 1] = dataTableList.Rows[j][k];
                            }
                            row++;
                        }
                    }
                }
            }
        }


        // Sayfa adı uzun ise ilk 15 karakteri alıyor.
        String sheetnamelenght(String value)
        {
            String control;
            if (value.Length < 30)
            {
                control = value;
            }
            else
            {
                control = value.Substring(0, 15).ToString();
            }
            return control;
        }


        // Kaydedilecek yeni excelde sayfa sayfa gezerek sütuna göre farklı olan satırların arasına boşluk bırakıp yeni excele aktarır.
        public void sheetRowSpace()
        {
            String sheetname = "";
            int sayac;
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                worksheet = workbook.Worksheets[i];
                defaultValue();
                sayac = 0;
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
                    default:
                        columncontrolnumber = 8;
                        break;
                }

                String cellvalue = "";
                for (int j = 1; j < rowsCount; j++)
                {
                    cellvalue = getReadCell(j, columncontrolnumber);

                    if (cellvalue == getReadCell(j - 1, columncontrolnumber) || cellvalue == ""
                        || getReadCell(j - 1, columncontrolnumber) == "")
                    {

                    }
                    else
                    {
                        dataTableList.Rows.InsertAt(emptyRowSpace(), sayac);
                        sayac++;
                    }
                    sayac++;
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
            MessageBox.Show("Kayıt işlemi tamamlanmıştır.");
        }


        // Açık olan exceli kapatır.
        public void excelquit()
        {
            workbook.Close();
            excel.Quit();
        }
    }
}
