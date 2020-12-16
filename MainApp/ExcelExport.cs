using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System;
using System.Collections.Generic;
using Octokit;
using System.Reflection;

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

        public void GenerateExcel<T>(List<Issue> models)
        {
            var dataSet = new DataSet();

            dataSet.Tables.Add(ConvertToDataTable<T>(models));

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

        DataTable ConvertToDataTable<T>(List<Issue> models)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Loop through all the properties            
            // Adding Column to our datatable
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }
            dataTable.Columns.Add("Repository Name");

            // Adding Row
            foreach (Issue item in models)
            {
                var values = new object[Props.Length + 1];
                int i;
                for (i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows  
                    values[i] = Props[i].GetValue(item, null);
                }

                values[i] = item.GetName();

                // Finally add value to datatable  
                dataTable.Rows.Add(values);
            }

            RemoveColumns(dataTable);

            return dataTable;
        }

        private void RemoveColumns(DataTable dataTable)
        {
            dataTable.Columns.Remove("CommentsUrl");
            dataTable.Columns.Remove("EventsUrl");
            dataTable.Columns.Remove("ClosedBy");
            dataTable.Columns.Remove("User");
            dataTable.Columns.Remove("Labels");
            dataTable.Columns.Remove("Assignee");
            dataTable.Columns.Remove("Assignees");
            dataTable.Columns.Remove("Milestone");
            dataTable.Columns.Remove("Comments");
            dataTable.Columns.Remove("PullRequest");
            dataTable.Columns.Remove("ClosedAt");
            dataTable.Columns.Remove("CreatedAt");
            dataTable.Columns.Remove("UpdatedAt");
            dataTable.Columns.Remove("Locked");
            dataTable.Columns.Remove("Repository");
            dataTable.Columns.Remove("Reactions");
        }
    }
}
