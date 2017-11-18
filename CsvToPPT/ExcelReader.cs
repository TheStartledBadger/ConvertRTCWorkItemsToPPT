using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace CsvToPPT
{
    public class ExcelReader
    {
        private string filename;

        public ExcelReader(string filename)
        {
            this.filename = filename;
        }

        public List<WorkItemInfo> getWorkItems()
        {
            List<WorkItemInfo> results = new List<WorkItemInfo>();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                string id = getCellValue(xlRange, i, 1);
                string summary = getCellValue(xlRange, i, 2);
                string owner = getCellValue(xlRange, i, 3);
                string costString = getCellValue(xlRange, i, 4);
                string plannedFor = getCellValue(xlRange, i, 5);
                int cost = Convert.ToInt32(costString);

                results.Add(new WorkItemInfo(id, summary, owner, cost, plannedFor));
            }

            xlWorkbook.Close();
            xlApp.Quit();
            return results;
        }

        private string getCellValue(Excel.Range xlRange, int row, int col)
        {
            if (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null)
            {
                return xlRange.Cells[row, col].Value2.ToString();
            }
            else
            {
                return "<undefined>";
            }
        }
    }
}