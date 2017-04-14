using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Test_Excel
{
    public static class ExcelHandler
    {

        public static void WriteFile(string path, DataTable data)
        {
            Exception ex = dataValidation(path, data);
            if (ex != null) throw ex;
            processExcel(path, data, File.Exists(path));
        }

        public static DataTable ReadFile(string path)
        {
            Exception ex = fileValidation(path);
            if (ex != null) throw ex;
            return readExcel(path);
        }

        private static Exception dataValidation(string path, DataTable data)
        {
            Exception ex = fileValidation(path);

            return (ex == null) ? (data == null || data.Rows.Count == 0) ? 
            new Exception("No data provided.") : null : ex;
        }

        private static  Exception fileValidation(string path)
        {
            return (string.IsNullOrEmpty(path)) ? new Exception("Invalid path provided.") : null;
        }

        /// <summary>
        /// Processes the excel file.
        /// </summary>
        /// <param name="path">Path were the file is or will be stored.</param>
        /// <param name="data">Datatable containing the data to be stored.</param>
        /// <param name="fileExsist">Flag to indicate if a new workbook will be created or overwritten.</param>
        private static void processExcel(string path, DataTable data, bool fileExsist)
        {
            ///////////////  Excel Initialization ///////////////
            Excel.Application app = new Excel.Application();
            Excel.Workbook wBook = (fileExsist) ? app.Workbooks.Open(path) : app.Workbooks.Add();
            Excel.Worksheet wSheet = wBook.Sheets[1];
            /////////////////////////////////////////////////////

            /////////////// Clear Previous Data  ///////////////
            Excel.Range wRange = wSheet.UsedRange;
            wRange.Clear();
            ////////////////////////////////////////////////////

            ///////////////  Actual Write to the file ///////////////
            writeData(data, wSheet);
            /////////////////////////////////////////////////////////

            ///////////////  Save the excel file with default values ///////////////
            wBook.SaveAs(path);
            ////////////////////////////////////////////////////////////////////////

            ///////////////  COM Objects disposal ///////////////
            Marshal.ReleaseComObject(wRange);
            Marshal.ReleaseComObject(wSheet);

            wBook.Close();
            Marshal.ReleaseComObject(wBook);

            app.Quit();
            Marshal.ReleaseComObject(app);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        
            /////////////////////////////////////////////////////
        }

        /// <summary>
        /// Writes data to an exceel worksheet
        /// </summary>
        /// <param name="data">Datatable containing the information.</param>
        /// <param name="sheet">Excel worksheet where the data will be stred.</param>
        private static  void writeData(DataTable data, Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int columnIndex = 1;

            ///Go through whole table by row/column.
            foreach(DataRow row in data.Rows)
            {
                columnIndex = 1;
                foreach(object column in row.ItemArray)
                {
                    sheet.Cells[rowIndex, columnIndex]= column.ToString();
                    columnIndex++;
                }
                rowIndex++;
            }

            //// Needed to release current work sheet since COM objects aree always passed as value and not by reference.
            Marshal.ReleaseComObject(sheet);
        }

        private static DataTable readExcel(string path)
        {
            ///////////////  Excel Initialization ///////////////
            Excel.Application app = new Excel.Application();
            Excel.Workbook wBook = app.Workbooks.Open(path);
            Excel.Worksheet wSheet = wBook.Sheets[1];
            /////////////////////////////////////////////////////

            /////////////// Get the used Range in Excel worksheet ///////////////
            Excel.Range wRange = wSheet.UsedRange;
            int rowCount = wRange.Rows.Count;
            int columnCount = wRange.Columns.Count;
            /////////////////////////////////////////////////////////////////////

            ///////////////  Get Datatable format ///////////////
            DataTable data = createDatatable(columnCount);
            /////////////////////////////////////////////////////

            /////////////// Perform the Excel worksheet read ///////////////
            for (int i = 1; i <= rowCount; i++)
            {
                DataRow row = data.NewRow();
                for(int j = 1; j <= columnCount; j++)
                {
                    //Excel base is 1 while row base is 0.
                    row[j - 1] = wRange.Cells[i, j];
                }
                data.Rows.Add(row);
            }
            ////////////////////////////////////////////////////////////////

            ///////////////  Save the excel file with default values ///////////////
            wBook.SaveAs(path);
            ////////////////////////////////////////////////////////////////////////

            ///////////////  COM Objects disposal ///////////////
            Marshal.ReleaseComObject(wRange);
            Marshal.ReleaseComObject(wSheet);

            wBook.Close();
            Marshal.ReleaseComObject(wBook);

            app.Quit();
            Marshal.ReleaseComObject(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            /////////////////////////////////////////////////////

            return data;
        }

        /// <summary>
        /// Creates an empry tadatable with the desired number of columns.
        /// </summary>
        /// <param name="columnCount">Number of columns to add.</param>
        /// <returns>Formatted datatable.</returns>
        static DataTable createDatatable(int columnCount)
        {
            DataTable tb = new DataTable();

            for (int i = 0; i < columnCount; i++)
            {
                tb.Columns.Add();
            }
            return tb;
        }
    }

   
}
