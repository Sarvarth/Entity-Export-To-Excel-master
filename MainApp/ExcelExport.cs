using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System;

namespace MainApp
{
    static class ExcelExport
    {
        private static Excel.Application excelApp = new Excel.Application();
        private static Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
        private static string _path = string.Empty;

        public static void GenerateExcel(DataTable dataTable, string path)
        {
            _path = path;

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // create a excel app along side with workbook and worksheet and give a name to it
            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                // add all the columns
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                // add all the rows
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }
            
        }

        public static void SaveAndCloseExcel()
        {
            Console.WriteLine("Saving your data");
            excelWorkBook.SaveAs(_path);
            
            Console.WriteLine("Closing!! Have a nice day");
            excelWorkBook.Close();
            excelApp.Quit();

            Console.ReadKey();
        }
    }
}
