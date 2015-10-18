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
    class Program
    {


        public void ExportToExcel(System.Data.DataTable dataTable)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;

            //Start Excel and get Application object
            excel = new Microsoft.Office.Interop.Excel.Application();

            //To make Excel visible
            excel.Visible = false;
            excel.DisplayAlerts = false;

            //Creation of a new workbook
            excelworkBook = excel.Workbooks.Add(Type.Missing);

            //Create a new Excel Worksheet

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = "Test Work Sheet";

            //resize columns????
            excelCellrange = excelSheet.Range[excelSheet.Cells[1,1], excelSheet.Cells[dataTable.Rows.Count,dataTable.Columns.Count]];
            excelCellrange.EntireColumn.AutoFit();

            //Microsoft.Office.Interop.Excel.Borders border = excelCellrange 
        }

        public static void runOperations() {
            //Open the Netsuite usage Worksheet
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(@"C:\test folder\NS - All Users\NSall-recovered.xlsx");
            Worksheet excelSheet = wb.ActiveSheet;

            //Get count of occupied rows
            int rowCount = excelSheet.UsedRange.Rows.Count;

            //Create a Data Table to store usage values
            System.Data.DataTable nsUsageTable = new System.Data.DataTable();
            //Create columns within nsUsageTable
            nsUsageTable.Columns.Add("User Name", typeof(string));
            nsUsageTable.Columns.Add("Total Count", typeof(int));
            nsUsageTable.Columns.Add("Leads Created", typeof(int));
            nsUsageTable.Columns.Add("Leads Updated", typeof(int));
            nsUsageTable.Columns.Add("Contacts Created", typeof(int));
            nsUsageTable.Columns.Add("Contacts Updated", typeof(int));
            nsUsageTable.Columns.Add("Opportunities Created", typeof(int));
            nsUsageTable.Columns.Add("Opportunities Updated", typeof(int));
            nsUsageTable.Columns.Add("Quotations Created", typeof(int));
            nsUsageTable.Columns.Add("Quotations Updated", typeof(int));


            //Loop through Excelsheet rows and add usernames to the data table

            for (int i = 1; i <= rowCount; i++)
            {
                //Create new data row
                DataRow dr = nsUsageTable.NewRow();

                dr["User Name"] = excelSheet.Cells[i, 1].Value.ToString();
                nsUsageTable.Rows.Add(dr);
            }


            //Close the Excel Workbook
            wb.Close();

            string filePath = @"C:\test folder\11th October 2015\";
            Operations operations = new Operations();
            nsUsageTable = operations.populateColumn(filePath + "New Leads.csv",
                                        "Leads Created",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "Updated Leads.csv",
                                        "Leads Updated",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "New Contacts.csv",
                                        "Contacts Created",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "Updated Contacts.csv",
                                        "Contacts Updated",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "New Opportunities.csv",
                                        "Opportunities Created",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "Updated Opportunities.csv",
                                        "Opportunities Updated",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "New Quotations.csv",
                                        "Quotations Created",
                                        nsUsageTable);

            nsUsageTable = operations.populateColumn(filePath + "Updated Quotations.csv",
                                        "Quotations Updated",
                                        nsUsageTable);

            //Add up activity count for each user
            /*
            foreach (DataRow row in nsUsageTable.Rows)
            {
                 int totalCount    = (int)row["Leads Created"] +
                                     (int)row["Leads Updated"] +
                                     (int)row["Contacts Created"] +
                                     (int)row["Contacts Updated"] +
                                     (int)row["Opportunities Created"] +
                                     (int)row["Opportunities Updated"] +
                                     (int)row["Quotations Created"] +
                                     (int)row["Quotations Updated"];

                 row["Total Count"] = totalCount;
            }
             */

            //Create a new Dataset
            DataSet dataSet = new DataSet("NetsuiteDetails");
            dataSet.Tables.Add(nsUsageTable);

            Console.Write("Operations successful");

            //Check if nsUsageTable has been populated with all the required data by displaying all the contained data
            foreach (DataRow row in nsUsageTable.Rows)
            {
                Console.WriteLine(row["User Name"].ToString());
            }



            Console.WriteLine("end");
            Console.ReadLine();
        }

        static void Main(string[] args)
        {

            runOperations();

            //Email email = new Email();
            //email.test();
        
        }

      

    }


    
}
