using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Excel_Ayırma
{
    public class ExcelIslemler
    {
        _Excel.Application excel = new _Excel.Application();
        _Excel.Workbook workbook;
        _Excel.Worksheet worksheet;
        Range range, range1, range2;
        int rowsCount = 0, columnsCount = 0;

        System.Data.DataTable dataTableList = new System.Data.DataTable("Excel-List");

        public String[] dizi = new String[10];
        int columncontrolnumber = 8;


        // Gelen adresdeki excel dosyasını açar ve dataTable methodunu çalıştırır. dataTable methodu çalıştıkdan sonra adresdeki exceli kapatır.
        public void excelOpen(String path)
        {
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[1];
            range = worksheet.UsedRange;
            rowsCount = range.Rows.Count;
            columnsCount = range.Columns.Count;
            dataTable();
            sheetnamelist();
            workbook.Close();
            excel.Quit();
        }


        // Exceldeki hücreyi getirir. Eğer hücre boş ise geriye "Empty" değerini dönderir.
        public String getReadCell(int i, int j)
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
                        dataTableList.Columns.Add(getReadCell(0, i), typeof(string));
                    }
                }

                String[] rows = new string[columnsCount];
                String cellvalue = "";
                int sayac = 0;

                for (int i = 1; i < rowsCount - 1; i++)
                {
                    cellvalue = getReadCell(i, columncontrolnumber);
                    for (int j = 0; j < columnsCount; j++)
                    {
                        rows[j] = getReadCell(i, j);
                    }
                    dataTableList.Rows.Add(rows);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Kayıtlar aktarılırken beklenmedik bir hata oluştu. " +
                    "Lütfen teknik birim ile iletişime geçiniz.\n Hata kodu:" + e.Message.ToString());
                excelquit();
            }
        }

        void sheetnamelist()
        {
            String cellvalue = "";
            int sayac = 0;
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
                        dizi[sayac] = sheetname;
                        sayac += 1;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

        }


        // dataTableList nesnesini farklı yerlerde kullanabilmek için oluşturuldu.
        public DataTable getDataTable() { return dataTableList; }


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


        // Excel dosyasına yeni sayfa ekler.
        public void addworksheet(String sheetname)
        {
            _Excel.Sheets sheets = workbook.Worksheets;
            if (!sheets.Equals(sheetname))
            {
                var xlYeniSayfa = (_Excel.Worksheet)sheets.Add();

                if (sheetname.Length < 30)
                {
                    xlYeniSayfa.Name = sheetname.ToString();
                }
                else
                {
                    xlYeniSayfa.Name = sheetname.Substring(0, 15).ToString();
                }
            }
        }


        // dataTableList nesnesini yeni bir excele kaydeder.
        public void saveExcel(String adres, String filename)
        {
            workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            for (int i = 0; i < dizi.Length; i++)
            {
                if (dizi[i] != null)
                {
                    addworksheet(dizi[i].ToString());
                    worksheet = workbook.Worksheets[1];
                    for (int j = 0; j < dataTableList.Columns.Count; j++)
                    {
                        range1 = (Range)worksheet.Cells[1, 1];
                        range1.Cells[1, j + 1] = dataTableList.Columns[j];
                    }
                }
                else
                    break;
            }
            int row = 1;
            String control = dizi[0];
            String sheetcellvalue = "";
            for (int i = 0; i < dataTableList.Rows.Count; i++)
            {
                for (int k = 0; k < dizi.Length; k++)
                {
                    if (dataTableList.Rows[i][columncontrolnumber].ToString() == dizi[k] && dizi[k] != null && dataTableList.Rows[i][columncontrolnumber].ToString() != null)
                    {
                        sheetcellvalue = dizi[k];
                        worksheet = workbook.Worksheets[sheetnamelenght(sheetcellvalue).ToString()];
                        break;
                    }
                }
                if (sheetcellvalue != control)
                {
                    control = sheetcellvalue;
                    row = 1;
                }
                for (int j = 0; j < dataTableList.Columns.Count; j++)
                {
                    range2 = (Range)worksheet.Cells[row, j + 1];
                    range2.Cells[2, 1] = dataTableList.Rows[i][j].ToString();
                }
                row++;
            }
            workbook.SaveAs(@adres + @"\" + filename, _Excel.XlFileFormat.xlWorkbookNormal);
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

        void sheetColumnControlNumber(String sheetname)
        {
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
                default:
                    columncontrolnumber = 8;
                    break;
            }

        }

        public void excelquit()
        {
            workbook.Close();
            excel.Quit();
        }
    }
}
