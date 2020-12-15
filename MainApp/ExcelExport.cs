using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System;

namespace MainApp
{
    class ExcelExport
    {
        private readonly string _path = string.Empty;
        private readonly Excel.Application excelApp;
        private readonly Excel.Workbook excelWorkBook;
        private readonly Excel.Worksheet excelWorkSheet;

        public ExcelExport(string path, string workSheetName)
        {
            _path = path;
            excelApp = new Excel.Application();
            excelWorkBook = excelApp.Workbooks.Add();

            //Add a new worksheet to workbook with the Datatable name
            excelWorkSheet = excelWorkBook.Sheets.Add();
            excelWorkSheet.Name = workSheetName;

        }

        public void GenerateExcel(DataTable dataTable)
        {
            var dataSet = new DataSet();

            dataSet.Tables.Add(dataTable);

            // create a excel app along side with workbook and worksheet and give a name to it
            foreach (DataTable table in dataSet.Tables)
            {
                
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

        public void SaveAndCloseExcel()
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
