using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;


namespace ExcelCSVImporter
{
    public partial class RibbonTools
    {
        private void RibbonTools_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void butReimport_Click(object sender, RibbonControlEventArgs e)
        {
            string filePath = "";
            Excel.Workbook currentWorkbok = Globals.ThisAddIn.Application.ActiveWorkbook;

            //if a file has been opened in Excel, the FullName will be the full path to the file.
            filePath = currentWorkbok.FullName;

            //testing
            //filePath = @"C:\Users\Chris\Desktop\test-csv.csv";
            //filePath = "";

            ProcessFile(filePath);
        }


        private void butImport_Click(object sender, RibbonControlEventArgs e)
        {
            string filePath = "";
            int cancelled = 0;
            object missing = System.Type.Missing;
            //object missing = System.Reflection.Obj;

            Microsoft.Office.Core.FileDialog fileDialog = Globals.ThisAddIn.Application.get_FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogOpen);
            fileDialog.AllowMultiSelect = false;
            fileDialog.Filters.Clear();            
            fileDialog.Filters.Add("CSV Files", "*.csv;*.txt", missing);

            //show dialog
            cancelled = fileDialog.Show();

            if (cancelled != 0)
            {
                filePath = fileDialog.SelectedItems.Item(1);
            }
            

            ProcessFile(filePath);
        }


        private void ProcessFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                System.Windows.Forms.MessageBox.Show("Not file specified.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }


            //simple try catch
            try
            {
                //timer for testing purposes
                var sw = new System.Diagnostics.Stopwatch();

                //start timer
                sw.Start();


                //disable screen updating during this process to increase performance
                Globals.ThisAddIn.Application.ScreenUpdating = false;


                //file data
                var fileRows = new List<string[]>();

                //need to set the Excel range
                //will compute  max columns as we loop through the file
                //this will accomodate "longer" lines                
                var maxNumberOfColumns = 0;


                //get data from file    
                //this should allow us to ready files already opened (needed for Reimport)
                using (System.IO.FileStream stream = System.IO.File.Open(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                {
                    //using VBs built-in csv parser
                    using (var reader = new Microsoft.VisualBasic.FileIO.TextFieldParser(stream))
                    {
                        reader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                        reader.SetDelimiters(",");

                        //read data
                        while (!reader.EndOfData)
                        {
                            var row = reader.ReadFields();

                            //add to list of rows
                            fileRows.Add(row);

                            //set max num of columns
                            if (row.Length > maxNumberOfColumns)
                            {
                                maxNumberOfColumns = row.Length;
                            }
                        }
                    }
                }                


                //excel range of data
                var rangeRowStart = 1;
                var rangeRowEnd = fileRows.Count;
                var rangeColStart = 1;
                var rangeColEnd = maxNumberOfColumns;


                //get current worksheet
                //this could be used to import dat into current worksheet
                //Excel.Worksheet currentWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

                //create new worksheet
                Excel.Worksheet newWorksheet;
                newWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Sheets.Add();

                //get range - entire range that will be used
                Excel.Range formatRange = newWorksheet.Range[newWorksheet.Cells[rangeRowStart, rangeColStart], newWorksheet.Cells[rangeRowEnd, rangeColEnd]];

                //set formatting for entire range
                formatRange.NumberFormat = "@";


                //array for output data (row X column)
                var rowIndex = 0;
                var colIndex = 0;
                var outputData = new string[fileRows.Count, rangeColEnd];


                //set output data
                for (rowIndex = 0; rowIndex < fileRows.Count; rowIndex++)
                {
                    var row = fileRows[rowIndex];
                    for (colIndex = 0; colIndex < row.Length; colIndex++)
                    {
                        outputData[rowIndex, colIndex] = row[colIndex];
                    }
                }


                //get data
                rowIndex = 0;
                colIndex = 0;
                //using (var reader = new Microsoft.VisualBasic.FileIO.TextFieldParser(filePath))
                //{
                //    reader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                //    reader.SetDelimiters(",");

                //    while (!reader.EndOfData)
                //    {
                //        var row = reader.ReadFields();

                //        for (colIndex = 0; colIndex < row.Length; colIndex++)
                //        {
                //            if (colIndex >= rangeColEnd)
                //            {
                //                //using header column to set max columns
                //                //if another row has more columns, we need to exit or set max columns by looping through entire data first                        
                //                break;
                //            }

                //            outputData[rowIndex, colIndex] = row[colIndex];
                //        }

                //        rowIndex++;
                //    }
                //}


                rowIndex = 0;
                colIndex = 0;
                //for (rowIndex = 0; rowIndex < fileLines.Length; rowIndex++)
                //{
                //    var items = Get_CSVItems(fileLines[rowIndex]);
                //    for (colIndex = 0; colIndex < items.Length; colIndex++)
                //    {
                //        if (colIndex >= rangeColEnd)
                //        {
                //            //using header column to set max columns
                //            //if another row has more columns, we need to exit or set max columns by looping through entire data first                        
                //            break;
                //        }

                //        outputData[rowIndex, colIndex] = items[colIndex];
                //    }
                //}


                //set ranges data
                formatRange.Value2 = outputData;


                //stop timer
                sw.Stop();
                //System.Diagnostics.Debug.WriteLine(sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);                
            }

            //update screen
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }


        //dumb parser
        private string[] Get_CSVItems(string row)
        {
            if (!string.IsNullOrWhiteSpace(row))
            {
                return row.Split(new char[] { ',' }, StringSplitOptions.None);
            }

            return new string[0];
        }
    }
}
