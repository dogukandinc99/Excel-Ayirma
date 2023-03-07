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
        _Excel.Application excel2 = new _Excel.Application();

        _Excel.Workbook workbook, workbook2;
        _Excel.Worksheet worksheet, worksheet2;
        Range range, range1, range2;
        int rowsCount = 0, columnsCount = 0;

        System.Data.DataTable dataTablelist = new System.Data.DataTable("Excel-List");

        public String[] dizi = new String[10];
        int columncontrolnumber = 3;


        // Gelen adresdeki excel dosyasını açar ve dataTable methodunu çalıştırır. dataTable methodu çalıştıkdan sonra adresdeki exceli kapatır.
        public void excelOpen(String path)
        {
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[1];
            range = worksheet.UsedRange;
            rowsCount = range.Rows.Count;
            columnsCount = range.Columns.Count;
            dataTable();
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


        // Adresdeki exceli dataTablelist nesnesine aktarır.
        void dataTable()
        {
            try
            {
                for (int i = 0; i < columnsCount; i++)
                {
                    dataTablelist.Columns.Add(getReadCell(0, i), typeof(string));
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
                    dataTablelist.Rows.Add(rows);
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
                MessageBox.Show("Kayıtlar aktarılırken beklenmedik bir hata oluştu. " +
                    "Lütfen teknik birim ile iletişime geçiniz.\n Hata kodu:" + e.Message.ToString());
                excelquit();
            }
        }


        // dataTablelist nesnesini farklı yerlerde kullanabilmek için oluşturuldu.
        public DataTable getDataTable() { return dataTablelist; }


        // Boş satır oluşturur.
        DataRow emptyRowSpace()
        {
            DataRow dr = dataTablelist.NewRow();
            for (int j = 0; j < columnsCount; j++)
            {
                dr[j] = "";
            }
            return dr;
        }


        // Excel dosyasına yeni sayfa ekler.
        public void addworksheet(String sheetname)
        {
            _Excel.Sheets sheets = workbook2.Worksheets;
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


        // dataTableList nesnesini yeni bir excele kaydeder.
        public void saveExcel(DataGridView dataGridView, String adres, String safefilename)
        {
            excel2.Visible = true;
            workbook2 = excel2.Workbooks.Add(System.Reflection.Missing.Value);
            for (int i = 0; i < dizi.Length; i++)
            {
                if (dizi[i] != null)
                {
                    addworksheet(dizi[i].ToString());
                }
                else
                    break;
            }
            worksheet2 = workbook2.Worksheets[1];
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                range1 = (Range)worksheet2.Cells[1, 1];
                range1.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;
            }
            int row = 1;
            String control = dizi[0];
            String sheetcellvalue = "";
            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
            {
                for (int k = 0; k < dizi.Length; k++)
                {
                    if (dataGridView[columncontrolnumber, i].Value.ToString() == dizi[k] && dizi[k] != null && dataGridView[columncontrolnumber, i].Value.ToString() != null)
                    {
                        sheetcellvalue = dizi[k];
                        worksheet2 = workbook2.Worksheets[sheetnamelenght(sheetcellvalue).ToString()];
                        break;
                    }
                }
                if (sheetcellvalue != control)
                {
                    control = sheetcellvalue;
                    row = 1;
                }
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    range2 = (Range)worksheet2.Cells[row, j + 1];
                    range2.Cells[2, 1] = dataGridView[j, i].Value;
                }
                row++;
            }
            workbook2.SaveAs(@adres + @"\" + "01-17_" + safefilename, _Excel.XlFileFormat.xlWorkbookNormal);
            excelquit();
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

        public void sheetSort(int controlcolumn)
        {
            String control = "";
            for (int i = 0; i < dizi.Length; i++)
            {
                if (dizi[i] != null)
                {
                    control = sheetnamelenght(dizi[i]);
                    worksheet2 = workbook2.Worksheets[control.ToString()];
                    int deger = int.Parse(Interaction.InputBox(control + " sayfası için kaçıncı sütuna bakılsın!", "Sütun Kontrol...").ToString());

                }
            }
        }

        public void excelquit()
        {
            /*excel2.Quit();
            excel2 = null;
            workbook2 = null;*/
        }
    }
}
