using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    public class Operations
    {
        public Operations()
        { }

        public void ExportDataSetToExcel(DataSet dataSet)
        { 
           
            //Create an Excel application instance
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:/test folder/NSdata.xlsx");

            foreach (System.Data.DataTable table in dataSet.Tables)
            { 
                //Add a new worksheet to the workbook with the Datatable's name
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = excelWorkbook.Sheets.Add();
                excelWorksheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorksheet.Cells[1, i] = table.Columns[i - 1];
                }
            }
        }


        public System.Data.DataTable populateColumn(string fileName,string columnToCheck,System.Data.DataTable nsUsageTable)
        {


            //Begin Leads Created Section

            //Read from the Leads Created CSV
            var reader = new StreamReader(File.OpenRead(fileName));
            List<string> names = new List<string>();
            List<string> usagevalues = new List<string>();

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');

                names.Add(values[0]);
                usagevalues.Add(values[1]);
            }//End of Reading from Leads CSV

            //Check the Data Table for matches and populate the leads column
            foreach (DataRow row in nsUsageTable.Rows)
            {
                for (int i = 0; i < names.Count; i++)
                {
                    if (row["User Name"].ToString() == names[i].ToString())
                    {
                        row[columnToCheck] = usagevalues[i];

                    }
                }
            }

            //End of Leads Created Section



            return nsUsageTable;
        }


    }
    
}
