using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace HelpNoteWorker
{
    class ExcelProgram
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Ensure all Excel documents are saved and closed. Press enter to continue.");
            Console.ReadLine();
            string path = "C:\\Users\\americaj\\Documents\\file.xlsx";
            string s = null;
            bool success = ProcessFile(path);
            if (success)
            {
                Process[] process = Process.GetProcesses();
                foreach (Process proc in process)
                {
                    s = proc.ProcessName.ToLower();
                    if (s.CompareTo("excel") == 0)
                    {
                        proc.Kill();
                    }
                }

                Console.WriteLine("Finished reformatting help documents.");
            }
            else
                Console.WriteLine("No files found.");
        }


        private static bool ProcessFile(string path)
        {
            bool success = true;
            try
            {
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Application app2 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
                Workbook xlWorkBook = app2.Workbooks.Add(missing);
                Worksheet xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);
                Worksheet sheet = (Worksheet)wb.Worksheets.get_Item(1);
                int row = 1;
                int startRow = 1;
                //int endRow = 10;
                for (int endRow = 10; endRow < 770; )
                {
                    Range excelRange = sheet.get_Range("C" + startRow.ToString(), "C" + endRow.ToString());
                    
                    System.Array myVals = (System.Array)excelRange.Cells.Value;
                    string[] strArray = new string[myVals.Length];
                    strArray = ConvertToStringArray(myVals);
                    int i = 1;
                    foreach (String item in strArray)
                    {
                        xlWorkSheet.Cells[row, i] = item;
                            i++;
                    }
                    endRow += 10;
                    startRow += 10;
                    row++;
                }



                xlWorkBook.SaveAs("C:\\Users\\americaj\\Documents\\JacobAmericano.xlsx");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                success = false;
            }
            return success;
        }

        public string[] GetRange(string range, Worksheet excelWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range workingRangeCells =
              excelWorksheet.get_Range(range, Type.Missing);
            //workingRangeCells.Select();

            System.Array array = (System.Array)workingRangeCells.Cells.Value2;
            string[] arrayS = ConvertToStringArray(array);

            return arrayS;
        }

        static string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(i, 1) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(i, 1).ToString();
            }

            return theArray;
        }

    }
}
