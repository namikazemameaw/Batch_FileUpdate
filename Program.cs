using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using static System.Net.WebRequestMethods;
using System.Threading;
using System.Runtime.Remoting.Lifetime;
using File = System.IO.File;
using System.Configuration;
using ClassLibrary;
using System.Reflection;

namespace Batch_FileUpdate
{
    internal class Program
    {
        static Log _log = new Log(MethodBase.GetCurrentMethod().DeclaringType.Namespace, MethodBase.GetCurrentMethod().Name, ConfigurationManager.AppSettings["PathLog"], ConfigurationManager.AppSettings["FirstNamefile"]);
        static void Main(string[] args)
        {

            _log.Info("start");

            Console.WriteLine("Start");
            //Command line argument
            if (args.Length != 0)
            {
                //Get Path 
                string[] subfiles = Directory.GetFiles(args[0], "*.*", SearchOption.AllDirectories);

                //Create DataTable
                DataTable dataTable = CreateDataTable();

                //Split Path in Array and Input Path File in DataTable
                DataTable resultDataTable = SplitePath(subfiles, dataTable);

                //Create Excel
                CreateExcel(resultDataTable);

            }
            //Enter your input
            if (args.Length == 0)
            {
                _log.Info("Please enter a file or directory path:");
                Console.WriteLine("Please enter a file or directory path:");
                string input = Console.ReadLine();

                //Check Input
                //// Input file or directory
                if (File.Exists(input) || Directory.Exists(input) )
                {
                    //Get Path 
                    string[] subfiles = Directory.GetFiles(input, "*.*", SearchOption.AllDirectories);

                    //Create DataTable
                    DataTable dataTable = CreateDataTable();

                    //Split Path in Array and Input Path File in DataTable
                    DataTable resultDataTable = SplitePath(subfiles, dataTable);
                    
                    //Create Excel
                    CreateExcel(resultDataTable);
                }
                ////input etc.
                else
                {
                    _log.Info("Input is not a valid file or directory path.");
                    Console.WriteLine("Input is not a valid file or directory path.");
                }


            }

        }
        private static DataTable CreateDataTable()
        {
            _log.Info("Start CreateDataTable");
            //CreateDataTable
            Console.WriteLine("Start CreateDataTable");
            DataTable dataTable = new DataTable();
            try { 

            //ColumnName
            DataColumn col1 = new DataColumn("No.");
            DataColumn col2 = new DataColumn("FileDeploy");
            DataColumn col3 = new DataColumn("SubFileDeploy");
            DataColumn col4 = new DataColumn("FileSize");

            //AddColumnName
            dataTable.Columns.Add(col1);
            dataTable.Columns.Add(col2);
            dataTable.Columns.Add(col3);
            dataTable.Columns.Add(col4);

            }
            catch (Exception ex)
            {
                _log.Info(ex.ToString());

                Console.WriteLine(ex);

            }
            _log.Info("End CreateDataTable");
            Console.WriteLine("End CreateDataTable");
            return dataTable;           
        }
        private static DataTable SplitePath(string[] subfiles, DataTable dataTable)
        {
            _log.Info("Start SplitePath");
            Console.WriteLine("Start SplitePath");
            try
            {
                //Splite PathAll
                for (int i = 0; i < subfiles.Length; i++)
                {
                    int count = 0;
                    string subdirectories = null;
                    string[] CheckDirectories = null;
                    string[] directories = subfiles[i].Split(Path.DirectorySeparatorChar);

                    //Check DirectoryName 
                    ////before
                    if (i != 0)
                    {
                        CheckDirectories = subfiles[i - 1].Split(Path.DirectorySeparatorChar);
                    }
                    ////after
                    else
                    {
                        CheckDirectories = subfiles[i].Split(Path.DirectorySeparatorChar);
                    }
                    //Splite Directory
                    
                    for (int j = 0; j < directories.Length; j++)
                    {
                        if (j == 3)
                        {
                            subdirectories = directories[j];

                            for (int k = j + 1; k < (directories.Length); k++)
                            {
                                subdirectories = subdirectories + "\\" + directories[k];
                            }
                        }

                    }

                    long fileSize = new FileInfo(subfiles[i]).Length;

                    //Value in DataTable
                    
                    if (i == 0)
                    {
                        count++;
                        dataTable.Rows.Add(count, subdirectories, subdirectories, fileSize);
                    }
                    else if (directories[3] != CheckDirectories[3])
                    {
                        count++;
                        dataTable.Rows.Add(count, subdirectories, subdirectories, fileSize);
                    }
                    else
                    {
                        dataTable.Rows.Add(" ", "", subdirectories, fileSize);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Info(ex.ToString());
                Console.WriteLine(ex);
            }
            _log.Info("End SplitePath");
            Console.WriteLine("End SplitePath");
            return dataTable;
            
        }
        private static void CreateExcel(DataTable resultDataTable)                                                                                                      
        {

            _log.Info("Start CreateExcel");
            Console.WriteLine("Start CreateExcel");
            try
            {
                // Create a new Excel application
                Application excel = new Application();
                excel.Visible = false;

                // Create a new workbook
                Workbook workbook = excel.Workbooks.Add(Type.Missing);

                // Create a new worksheet
                Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

                worksheet.Cells[2, 1] = "Type :";
                worksheet.Cells[2, 1].ColumnWidth = 20;
                worksheet.Cells[2, 1].Style.Font.Bold = true;

                worksheet.Cells[2, 2].ColumnWidth = 20;
                worksheet.Cells[3, 1] = "Project :";

                //worksheet.Cells[3, 2] = folderName;

                // Set the column headers
                for (int i = 0; i < resultDataTable.Columns.Count; i++)
                {
                    worksheet.Cells[4, i + 3] = resultDataTable.Columns[i].ColumnName;
                    worksheet.Cells[4, i + 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    worksheet.Cells[4, 4].ColumnWidth = 40;
                    worksheet.Cells[4, 5].ColumnWidth = 50;
                    worksheet.Cells[4, i + 3].Borders.Weight = 1d;

                    // Set the cell values
                    for (int row = 0; row < resultDataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < resultDataTable.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 5, col + 3].Borders.Weight = 1d;
                            worksheet.Cells[row + 5, col + 3] = resultDataTable.Rows[row][col].ToString();
                            worksheet.Cells[row + 5, col + 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }
                    }
                }
                // Save the workbook and close Excel
                workbook.SaveAs(ConfigurationManager.AppSettings["Path"]);
                excel.Quit();
                _log.Info("End");
                Console.WriteLine("End");
            }
            catch (Exception ex)
            {
                _log.Info(ex.ToString());
                Console.WriteLine(ex);
            }
        }
    }
}
