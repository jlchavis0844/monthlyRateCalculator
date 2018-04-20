using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Linq;

namespace monthlyRateCalculator {
    class Program {


        [STAThread]
        static void Main(string[] args) {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select 12 month folder";
            fbd.SelectedPath = @"\\nas3\Shared\RALIM\TDSGroup-Kronos\Allstate 2018\Plan Setups\";
            //fbd.Filter = "Excel Files | *.xls";
            //fbd.FilterIndex = 2;
            //fbd.RestoreDirectory = true;

            if(fbd.ShowDialog() == DialogResult.OK) {
                if (Directory.Exists(fbd.SelectedPath)) {
                    string filePath = fbd.SelectedPath;
                    string rootPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath),@".\"));
                    string[] files = Directory.GetFiles(filePath, @"*.xls");

                    foreach (string file in files) {
                        int numOfMon = 11;
                        string fileName = Path.GetFileName(file);
                        List<string> paths = new List<string>();
                        //paths.Add(rootPath + @"\12-mon\");
                        paths.Add(rootPath + @"11-mon\");
                        paths.Add(rootPath + @"10-mon\");
                        paths.Add(rootPath + @"9-mon\");

                        foreach (string path in paths) {
                            Excel.Application xlApp;
                            Excel.Workbook xlWorkBook;
                            Excel.Worksheet xlWorkSheet;

                            object misValue = System.Reflection.Missing.Value;
                            xlApp = new Excel.Application();
                            xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true,
                                Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                            int nColumns = xlWorkSheet.UsedRange.Columns.Count;
                            int nRows = xlWorkSheet.UsedRange.Rows.Count;

                            for (int row = 2; row <= nRows; row++) {
                                double tempAnswer = Convert.ToDouble(xlWorkSheet.Cells[5][row].Value2) * 12 / numOfMon;
                                Console.WriteLine("Changing " + xlWorkSheet.Cells[5][row].Value2 + " to " + tempAnswer);
                                xlWorkSheet.Cells[5][row].Value2 = Convert.ToDouble(xlWorkSheet.Cells[5][row].Value2) * 12 / numOfMon;
                            }

                            Console.WriteLine("Done!");
                            new FileInfo(path).Directory.Create();
                            string myFile = path + fileName;
                            xlWorkBook.SaveAs(myFile,
                                56, //Seems to work better than default excel 16
                                Type.Missing,
                                Type.Missing,
                                false,
                                false,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing,
                                Type.Missing,
                                Type.Missing,
                                Type.Missing,
                                Type.Missing);
                            xlWorkBook.Close();
                            xlApp.Quit();
                            releaseObject(xlWorkBook);
                            releaseObject(xlApp);

                            numOfMon--;
                        }
                    }
                }
            }
        }

        private static void releaseObject(object obj) {
            try {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex) {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally {
                GC.Collect();
            }
        }
    }
}
