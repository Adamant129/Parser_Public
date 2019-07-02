using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using ConsoleTestResultsParser.Models;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ConsoleTestResultsParser
{
    public static class FileWriter
    {
        public static HSSFWorkbook hssfworkbook = new HSSFWorkbook();
        public static ICellStyle DefaultHeaderCellStyle;
        public static ICellStyle DefaultCellStyle;
        public static ICellStyle PassedTestCellStyle;
        public static ICellStyle FailedTestCellStyle;
        public static ICellStyle SkippedTestCellStyle;

        public static void WriteXls(List<TestRun> runs, string[] args)
        {
            var currentDate = DateTime.Now.ToString("M-d-yyyy");
            string path;
            if (args.Length == 0) path = Path.Combine(ConfigurationManager.AppSettings["xlsResultFileFolder"], currentDate + "_Daily_Tests_Run.xls");
            else path = Path.Combine(args[0], currentDate + "_Daily_Tests_Run.xls");

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "DFS Venue Regression Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            CreateCellStyles();

            CreateAllRunsDataTab(runs);
            CreateEnvironmentsResultsTab(runs);
            CreateFailedTestsTab(runs);


            FileStream file = new FileStream(path, FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        private static void CreateAllRunsDataTab(List<TestRun> runs)
        {
            int totalCellsAmount = 9;
            int rowNumberIterator = 0;

            var defaultSheet = hssfworkbook.CreateSheet("All runs data");
            defaultSheet.SetColumnWidth(0, 12 * 270);
            defaultSheet.SetColumnWidth(1, 25 * 270);
            defaultSheet.SetColumnWidth(2, 10 * 270);
            defaultSheet.SetColumnWidth(3, 20 * 270);
            defaultSheet.SetColumnWidth(4, 20 * 270);
            defaultSheet.SetColumnWidth(5, 12 * 270);
            defaultSheet.SetColumnWidth(6, 14 * 270);
            defaultSheet.SetColumnWidth(7, 20 * 270);
            defaultSheet.SetColumnWidth(8, 20 * 270);
            defaultSheet.SetColumnWidth(9, 20 * 270);
            defaultSheet.SetColumnWidth(10, 20 * 270);

            var firstRow = defaultSheet.CreateRow(0);
            firstRow.CreateCell(0).SetCellValue("Environment");
            firstRow.CreateCell(1).SetCellValue("Venue version");
            firstRow.CreateCell(2).SetCellValue("Run Result");
            firstRow.CreateCell(3).SetCellValue("Start Time");
            //firstRow.CreateCell(3).SetCellValue("End Time"); //Removed, required by Maksim
            firstRow.CreateCell(4).SetCellValue("Duration (min)");
            firstRow.CreateCell(5).SetCellValue("Number Of Tests");
            firstRow.CreateCell(6).SetCellValue("Number Of Passed Tests");
            firstRow.CreateCell(7).SetCellValue("Number Of Failed Tests");
            firstRow.CreateCell(8).SetCellValue("Number Of Skipped Tests");
            //firstRow.CreateCell(9).SetCellValue("Change since Yesterday"); //Removed, required by Maksim

            for (int p = 0; p < totalCellsAmount; p++)
            {
                firstRow.GetCell(p).CellStyle = DefaultHeaderCellStyle;
            }

            foreach (var run in runs)
            {
                ICellStyle regularStyle = hssfworkbook.CreateCellStyle();
                regularStyle.Alignment = HorizontalAlignment.Center;
                regularStyle.VerticalAlignment = VerticalAlignment.Center;
                regularStyle.FillPattern = FillPattern.SolidForeground;

                var sheetSecondRow = defaultSheet.CreateRow(rowNumberIterator + 1);
                sheetSecondRow.CreateCell(0).SetCellValue(run.Environment);
                sheetSecondRow.CreateCell(1).SetCellValue(run.VenueVersion);
                sheetSecondRow.CreateCell(2).SetCellValue(run.RunResult);
                sheetSecondRow.CreateCell(3).SetCellValue(run.StartTime.ToString());
                //secondRow.CreateCell(3).SetCellValue(run.EndTime.ToString()); //Removed, required by Maksim
                sheetSecondRow.CreateCell(4).SetCellValue(run.Duration.ToString(@"hh\:mm\:ss"));
                sheetSecondRow.CreateCell(5).SetCellValue(run.NumberOfTests);
                sheetSecondRow.CreateCell(6).SetCellValue(run.NumberOfPassedTests);
                sheetSecondRow.CreateCell(7).SetCellValue(run.NumberOfFailedTests);
                sheetSecondRow.CreateCell(8).SetCellValue(run.NumberOfSkippedTests);
                //secondRow.CreateCell(9).SetCellValue("No changes"); //Removed, required by Maksim

                for (int p = 0; p < totalCellsAmount; p++)
                {
                    switch (run.RunResult)
                    {
                        case "Failed":
                            sheetSecondRow.GetCell(p).CellStyle = FailedTestCellStyle;
                            break;
                        case "Passed":
                            sheetSecondRow.GetCell(p).CellStyle = PassedTestCellStyle;
                            break;
                        default:
                            sheetSecondRow.GetCell(p).CellStyle = DefaultCellStyle;
                            break;
                    }
                }
                rowNumberIterator++;

                for (int p = 0; p < totalCellsAmount; p++)
                {
                    defaultSheet.AutoSizeColumn(p);
                }
            }
        }

        private static void CreateEnvironmentsResultsTab(List<TestRun> runs)
        {
            foreach (var run in runs)
            {
                var totalExecutionTime = 0.0;
                int totalCellsAmount = 7;
                try
                {
                    var sheet = hssfworkbook.CreateSheet(run.FileName.Split('\\').Last());
                    var sheetFirstRow = sheet.CreateRow(0);

                    sheetFirstRow.CreateCell(0).SetCellValue("#");
                    sheetFirstRow.CreateCell(1).SetCellValue("Test Name");
                    sheetFirstRow.CreateCell(2).SetCellValue("Test Result");
                    sheetFirstRow.CreateCell(3).SetCellValue("Test Duration (sec)");
                    sheetFirstRow.CreateCell(4).SetCellValue("Server ID");
                    sheetFirstRow.CreateCell(5).SetCellValue("UAE");
                    sheetFirstRow.CreateCell(6).SetCellValue("Reason");

                    for (int p = 0; p < totalCellsAmount; p++)
                    {
                        sheetFirstRow.GetCell(p).CellStyle = DefaultHeaderCellStyle;
                    }

                    foreach (var test in run.TestsInRun)
                    {
                        totalExecutionTime += test.Duration.TotalSeconds;

                        var sheetSecondRow = sheet.CreateRow(test.Number);
                        sheetSecondRow.CreateCell(0).SetCellValue(test.Number);
                        sheetSecondRow.CreateCell(1).SetCellValue(test.Name);
                        sheetSecondRow.CreateCell(2).SetCellValue(test.Result);
                        sheetSecondRow.CreateCell(3).SetCellValue(test.Duration.TotalSeconds);
                        sheetSecondRow.CreateCell(4).SetCellValue(test.ServerId);

                        if (test.Result == "Failed")
                        {
                            sheetSecondRow.CreateCell(5).SetCellValue(test.UAE);
                        }
                        else
                        {
                            sheetSecondRow.CreateCell(5).SetCellValue("N/A");
                        }

                        sheetSecondRow.CreateCell(6).SetCellValue("N/A");

                        for (int p = 0; p < totalCellsAmount; p++)
                        {
                            switch (test.Result)
                            {
                                case "Failed":
                                    sheetSecondRow.GetCell(p).CellStyle = FailedTestCellStyle;
                                    break;
                                case "Passed":
                                    sheetSecondRow.GetCell(p).CellStyle = PassedTestCellStyle;
                                    break;
                                default:
                                    sheetSecondRow.GetCell(p).CellStyle = DefaultCellStyle;
                                    break;
                            }
                        }
                    }

                    for (int p = 0; p < totalCellsAmount; p++)
                    {
                        sheet.AutoSizeColumn(p);
                        //GC.Collect();
                    }

                    var sheetLastRow = sheet.CreateRow(run.TestsInRun.Count + 1);
                    sheetLastRow.CreateCell(2).SetCellValue("Total: ");
                    sheetLastRow.CreateCell(3).SetCellValue(totalExecutionTime);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine("Error while creating Sheet." + e.Message);
                }
            }
        }

        private static void CreateFailedTestsTab(List<TestRun> runs)
        {

            List<string> cells = new List<string>
            {
                "Failed test",
                "Environment",
                "Failed again",
            };


            var rowNumberIterator = 1;
            var addExceptionColumn = Convert.ToBoolean(ConfigurationManager.AppSettings["addExceptionColumn"]);

            var failedTestSheet = hssfworkbook.CreateSheet("Failed tests");
            var failedTestSheetFirstRow = failedTestSheet.CreateRow(0);

            if (addExceptionColumn)
            {
                cells.Add("Stack trace");
            }

            foreach (var cell in cells)
            {
                failedTestSheetFirstRow.CreateCell(cells.IndexOf(cell)).SetCellValue(cell);
                failedTestSheetFirstRow.GetCell(cells.IndexOf(cell)).CellStyle = DefaultHeaderCellStyle;
            }

            foreach (var run in runs)
            {
                foreach (var test in run.TestsInRun.Where(p => p.Result == "Failed"))
                {
                    var sheetSecondRow = failedTestSheet.CreateRow(rowNumberIterator);

                    sheetSecondRow.CreateCell(cells.IndexOf("Failed test")).SetCellValue(string.Concat(run.StartTime.ToShortDateString(), " ", test.Name));
                    sheetSecondRow.CreateCell(cells.IndexOf("Environment")).SetCellValue(run.Environment);
                    sheetSecondRow.CreateCell(cells.IndexOf("Failed again")).SetCellType(CellType.Boolean);

                    if (addExceptionColumn)
                        sheetSecondRow.CreateCell(cells.IndexOf("Stack trace")).SetCellValue(test.StackTrace);

                    rowNumberIterator++;

                    for (int p = 0; p < cells.Count; p++)
                    {
                        sheetSecondRow.GetCell(p).CellStyle = FailedTestCellStyle;
                    }
                }

                for (int p = 0; p < cells.Count; p++)
                {
                    failedTestSheet.AutoSizeColumn(p);
                    //GC.Collect();
                }
            }
        }

        private static void CreateCellStyles()
        {
            DefaultHeaderCellStyle = hssfworkbook.CreateCellStyle();
            DefaultHeaderCellStyle.Alignment = HorizontalAlignment.Center;
            DefaultHeaderCellStyle.VerticalAlignment = VerticalAlignment.Center;
            DefaultHeaderCellStyle.BorderBottom = BorderStyle.Medium;

            DefaultCellStyle = hssfworkbook.CreateCellStyle();
            DefaultCellStyle.Alignment = HorizontalAlignment.Center;
            DefaultCellStyle.VerticalAlignment = VerticalAlignment.Center;
            DefaultCellStyle.FillPattern = FillPattern.SolidForeground;
            DefaultCellStyle.FillForegroundColor = IndexedColors.White.Index;

            PassedTestCellStyle = hssfworkbook.CreateCellStyle();
            PassedTestCellStyle.Alignment = HorizontalAlignment.Center;
            PassedTestCellStyle.VerticalAlignment = VerticalAlignment.Center;
            PassedTestCellStyle.FillPattern = FillPattern.SolidForeground;
            PassedTestCellStyle.FillForegroundColor = IndexedColors.LightGreen.Index;

            FailedTestCellStyle = hssfworkbook.CreateCellStyle();
            FailedTestCellStyle.Alignment = HorizontalAlignment.Center;
            FailedTestCellStyle.VerticalAlignment = VerticalAlignment.Center;
            FailedTestCellStyle.FillPattern = FillPattern.SolidForeground;
            FailedTestCellStyle.FillForegroundColor = IndexedColors.Coral.Index;
        }
    }
}
