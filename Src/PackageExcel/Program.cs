

using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Threading.Tasks;

namespace PackageExcel
{
    class Program
    {
        private static Application app;

        static void Main(string[] args)
        {
            app = new Application();
            var str = File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\Setting.json");
            var setting = JsonConvert.DeserializeObject<Setting>(str);
            TaskFactory f = new TaskFactory();
            int index = 1;

            foreach (var id in setting.FileIDs)
            {
                Console.WriteLine();
                Console.WriteLine($"**********************    index: {index++},  id: {id}");
                Task.Run(() => process(setting.FilePath, id, setting.Folder)).GetAwaiter().GetResult();
            }

            app.Quit();
        }

        private static void process(string filePath, string fileId, string subFolder)
        {
            string[] frontPage = Directory.GetFiles(filePath, $"*{fileId}*封皮*.xls");
            if (frontPage.Length != 1)
            {
                Console.WriteLine($"There are {frontPage.Length} front pages for {fileId}.");
            }
            if (frontPage.Length == 0)
            {
                return;
            }

            string[] folders = Directory.GetDirectories(filePath, $"UBS*{fileId}*");
            if (folders.Length != 1)
            {
                Console.WriteLine($"There are {folders.Length} pages for {fileId}.");
            }
            if (folders.Length == 0)
            {
                return;
            }

            string[] page = Directory.GetFiles(folders[0], $"*{fileId}*.xls*");
            Array.Sort(page, new MyComparer());
            if (page.Length == 0)
            {
                Console.WriteLine($"There are {page.Length} pages for {fileId}.");
            }
            if (page.Length == 0)
            {
                return;
            }

            var ps = new string[page.Length + 1];
            ps[0] = frontPage[0];
            for (int i = 0; i < page.Length; i++)
            {
                ps[i + 1] = page[i];
            }

            Console.WriteLine($"   front page  :   {frontPage[0]}.");
            Console.WriteLine($"   page number :   {page.Length}");

            MergeWorkbooks(ps, filePath + subFolder + $"{fileId}\\");
        }

        private static void MergeWorkbooks(string[] sourceFilePaths, string newFolder)
        {
            app = new Application();
            app.DisplayAlerts = false; // No prompt when overriding

            // open source workbooks (index=2,3,...)
            foreach (var sourceFilePath in sourceFilePaths)
            {
                app.Workbooks.Add(sourceFilePath);
            }

            var covernames = sourceFilePaths[0].Split("\\".ToCharArray());
            var coverName = covernames[covernames.Length - 1];
            coverName = coverName.Split(".".ToCharArray())[0] + ".xlsx";

            //format workbook
            FormatWorkBook(app.Workbooks, coverName, newFolder);

            Directory.CreateDirectory(newFolder);


            // Save new workbook
            // Close source documents before saving destination. Otherwise, save will fail
            int wbIndex = 0;
            while (app.Workbooks.Count > 0)
            {
                Workbook wb = app.Workbooks[1];
                var names = sourceFilePaths[wbIndex++].Split("\\".ToCharArray());
                var n = names[names.Length - 1].Split(".".ToCharArray())[0] + ".xlsx";
                wb.SaveAs(newFolder + n);
                wb.Close();
            }

            app.Workbooks.Close();
        }

        private static void FormatWorkBook(Workbooks wbs, string coverName, string rootPath)
        {
            int TotalSheet = GetSheetCount(wbs);
            if (wbs.Count == 0)
            {
                return;
            }
            var cover = wbs[1];
            cover.Sheets[1].Cells[3, 24].Value = $"第   1   页    共   {TotalSheet}   页";

            Console.Write("Sheet: ");
            int i = 2;
            for (int wbIndex = 2; wbIndex <= wbs.Count; wbIndex++)
            {
                Workbook wb = wbs[wbIndex];

                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (IsNotNullSheet(sheet))
                    {
                        Console.Write($"  {i}");
                        if (IsTypeA(sheet))
                        {
                            Console.Write($"<A>");
                            ProcessTypeA(cover.Sheets[1], sheet, coverName, i, TotalSheet, rootPath);
                        }
                        else if (IsTypeB(sheet))
                        {
                            Console.Write($"<B>");
                            ProcessTypeB(cover.Sheets[1], sheet, coverName, i, TotalSheet, rootPath);
                        }
                        else if(IsTypeC(sheet))
                        {
                            Console.Write($"<C>");
                            ProcessTypeC(cover.Sheets[1], sheet, coverName, i, TotalSheet, rootPath);
                        }
                        else
                        {
                            Console.Write($"<None>");
                        }
                        i++;
                    }
                }
            }
            Console.WriteLine();

        }

        private static int GetSheetCount(Workbooks wbs)
        {
            int count = 0;
            foreach (Workbook wb in wbs)
            {
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (IsNotNullSheet(sheet))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        private static bool IsNotNullSheet(Worksheet sheet)
        {
            Pictures pics = sheet.Pictures(System.Type.Missing) as Pictures;
            return pics.Count > 0;
        }

        // with DOCUMENT NUMBER
        private static bool IsTypeA(Worksheet sheet)
        {
            var s = sheet.Cells[2, 22].Value as string;
            return s == "DOCUMENT NUMBER";
        }

        private static void ProcessTypeA(Worksheet coverSheet, Worksheet sheet, string coverName, int num, int TotalSheet, string rootPath)
        {
            sheet.Range[sheet.Cells[2, 22], sheet.Cells[2, 26]].UnMerge();
            sheet.Range[sheet.Cells[3, 22], sheet.Cells[3, 26]].UnMerge();
            sheet.Range[sheet.Cells[5, 22], sheet.Cells[6, 26]].UnMerge();

            sheet.Range[sheet.Cells[1, 20], sheet.Cells[2, 21]].Merge();
            coverSheet.Cells[1, 24].Copy(sheet.Range[sheet.Cells[1, 20], sheet.Cells[2, 21]]);
            sheet.Range[sheet.Cells[1, 22], sheet.Cells[2, 26]].Merge();
            coverSheet.Cells[1, 25].Copy(sheet.Range[sheet.Cells[1, 22], sheet.Cells[2, 26]]);
            sheet.Range[sheet.Cells[1, 22], sheet.Cells[2, 26]].Value = $"='{rootPath}[{coverName}]中文首页'!Y1";

            sheet.Range[sheet.Cells[1, 27], sheet.Cells[2, 29]].Merge();

            sheet.Range[sheet.Cells[3, 20], sheet.Cells[4, 21]].Merge();
            coverSheet.Cells[2, 24].Copy(sheet.Range[sheet.Cells[3, 20], sheet.Cells[4, 21]]);

            sheet.Range[sheet.Cells[3, 22], sheet.Cells[4, 26]].Merge();
            coverSheet.Cells[2, 25].Copy(sheet.Range[sheet.Cells[3, 22], sheet.Cells[4, 26]]);
            sheet.Range[sheet.Cells[3, 22], sheet.Cells[4, 26]].Value = $"='{rootPath}[{coverName}]中文首页'!Y2";
            sheet.Range[sheet.Cells[3, 27], sheet.Cells[4, 29]].Merge();
            coverSheet.Cells[2, 28].Copy(sheet.Range[sheet.Cells[3, 27], sheet.Cells[4, 29]]);
            sheet.Range[sheet.Cells[3, 27], sheet.Cells[4, 29]].Value = $"='{rootPath}[{coverName}]中文首页'!AB2";

            sheet.Range[sheet.Cells[5, 20], sheet.Cells[6, 29]].Merge();
            sheet.Range[sheet.Cells[5, 20], sheet.Cells[6, 29]].Value = $"第   {num}   页    共   {TotalSheet}   页"; ;

            sheet.Range[sheet.Cells[1, 20], sheet.Cells[6, 29]].Font.Name = "宋体";
            sheet.Range[sheet.Cells[1, 20], sheet.Cells[6, 29]].Font.Size = 9;
            sheet.Range[sheet.Cells[1, 20], sheet.Cells[6, 29]].Font.Bold = false;

            sheet.Range[sheet.Cells[1, 20], sheet.Cells[6, 29]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            var size = sheet.Columns[19].ColumnWidth;
            sheet.Columns[19].ColumnWidth = 0;
            sheet.Columns[20].ColumnWidth = sheet.Columns[20].ColumnWidth + size;

            size = 0.7;
            sheet.Columns[22].ColumnWidth = sheet.Columns[22].ColumnWidth - size;
            sheet.Columns[27].ColumnWidth = sheet.Columns[27].ColumnWidth + size;

            size = 0.5;
            sheet.Columns[22].ColumnWidth = sheet.Columns[22].ColumnWidth - size;
            sheet.Columns[20].ColumnWidth = sheet.Columns[20].ColumnWidth + size;

            size = 2.0;
            sheet.Columns[25].ColumnWidth = sheet.Columns[25].ColumnWidth - size;
            sheet.Columns[27].ColumnWidth = sheet.Columns[27].ColumnWidth + size;

            sheet.Range[sheet.Cells[1, 27], sheet.Cells[2, 29]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Range[sheet.Cells[1, 22], sheet.Cells[2, 26]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
            sheet.Range[sheet.Cells[1, 27], sheet.Cells[2, 29]].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Range[sheet.Cells[1, 22], sheet.Cells[2, 26]].Borders[XlBordersIndex.xlEdgeBottom].Weight;

            sheet.Range[sheet.Cells[3, 27], sheet.Cells[4, 29]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Range[sheet.Cells[3, 22], sheet.Cells[4, 26]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
            sheet.Range[sheet.Cells[3, 27], sheet.Cells[4, 29]].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Range[sheet.Cells[3, 22], sheet.Cells[4, 26]].Borders[XlBordersIndex.xlEdgeBottom].Weight;
        }


        // with UAN
        private static bool IsTypeB(Worksheet sheet)
        {
            var s = sheet.Cells[1, 23].Value as string;
            return s == "UAN";
        }

        private static void ProcessTypeB(Worksheet coverSheet, Worksheet sheet, string coverName, int num, int TotalSheet, string rootPath)
        {
            sheet.Range[sheet.Cells[1, 23], sheet.Cells[3, 30]].UnMerge();
            sheet.Range[sheet.Cells[3, 23], sheet.Cells[3, 30]].Merge();

            coverSheet.Cells[1, 24].Copy(sheet.Cells[1, 23]);
            coverSheet.Cells[1, 25].Copy(sheet.Cells[1, 25]);
            sheet.Range[sheet.Cells[1, 25], sheet.Cells[1, 25]].Value = $"='{rootPath}[{coverName}]中文首页'!Y1";

            coverSheet.Cells[2, 24].Copy(sheet.Cells[2, 23]);
            coverSheet.Cells[2, 25].Copy(sheet.Cells[2, 25]);
            sheet.Range[sheet.Cells[2, 25], sheet.Cells[2, 25]].Value = $"='{rootPath}[{coverName}]中文首页'!Y2";
            coverSheet.Cells[2, 28].Copy(sheet.Cells[2, 29]);
            sheet.Range[sheet.Cells[2, 29], sheet.Cells[2, 29]].Value = $"='{rootPath}[{coverName}]中文首页'!AB2";

            sheet.Range[sheet.Cells[3, 23], sheet.Cells[3, 30]].Value = $"第   {num}   页    共   {TotalSheet}   页";

            sheet.Range[sheet.Cells[1, 23], sheet.Cells[3, 30]].Font.Name = "宋体";
            sheet.Range[sheet.Cells[1, 23], sheet.Cells[3, 30]].Font.Size = 10;
            sheet.Range[sheet.Cells[1, 23], sheet.Cells[3, 30]].Font.Bold = false;

            sheet.Range[sheet.Cells[1, 23], sheet.Cells[3, 30]].Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlLineStyleNone;

            sheet.Range[sheet.Cells[1, 23], sheet.Cells[2, 30]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Range[sheet.Cells[2, 23], sheet.Cells[2, 23]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
            sheet.Range[sheet.Cells[1, 23], sheet.Cells[2, 30]].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Range[sheet.Cells[2, 23], sheet.Cells[2, 23]].Borders[XlBordersIndex.xlEdgeBottom].Weight;

            sheet.Range[sheet.Cells[1, 23], sheet.Cells[1, 30]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Range[sheet.Cells[2, 23], sheet.Cells[2, 23]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
            sheet.Range[sheet.Cells[1, 23], sheet.Cells[1, 30]].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Range[sheet.Cells[2, 23], sheet.Cells[2, 23]].Borders[XlBordersIndex.xlEdgeBottom].Weight;

        }

        // with DOCUMENT NUMBER
        private static bool IsTypeC(Worksheet sheet)
        {
            var s = sheet.Cells[1, 30].Value as string;
            return s == "DOCUMENT NUMBER";
        }

        private static void ProcessTypeC(Worksheet coverSheet, Worksheet sheet, string coverName, int num, int TotalSheet, string rootPath)
        {
            sheet.Range[sheet.Cells[1, 30], sheet.Cells[2, 39]].UnMerge();

            coverSheet.Cells[1, 24].Copy(sheet.Range[sheet.Cells[1, 30], sheet.Cells[1, 30]]);
            coverSheet.Cells[1, 25].Copy(sheet.Range[sheet.Cells[1, 33], sheet.Cells[1, 33]]);
            sheet.Range[sheet.Cells[1, 33], sheet.Cells[1, 33]].Value = $"='{rootPath}[{coverName}]中文首页'!Y1";

            coverSheet.Cells[2, 24].Copy(sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 30]]);
            coverSheet.Cells[2, 25].Copy(sheet.Range[sheet.Cells[2, 33], sheet.Cells[2, 33]]);
            sheet.Range[sheet.Cells[2, 33], sheet.Cells[2, 33]].Value = $"='{rootPath}[{coverName}]中文首页'!Y2";
            coverSheet.Cells[2, 28].Copy(sheet.Range[sheet.Cells[2, 37], sheet.Cells[2, 37]]);
            sheet.Range[sheet.Cells[2, 37], sheet.Cells[2, 37]].Value = $"='{rootPath}[{coverName}]中文首页'!AB2";

            sheet.Range[sheet.Cells[3, 30], sheet.Cells[3, 39]].Value = $"第   {num}   页    共   {TotalSheet}   页"; ;

            sheet.Range[sheet.Cells[1, 30], sheet.Cells[3, 29]].Font.Name = "宋体";
            sheet.Range[sheet.Cells[1, 30], sheet.Cells[3, 29]].Font.Size = 9;
            sheet.Range[sheet.Cells[1, 30], sheet.Cells[3, 29]].Font.Bold = false;

            sheet.Range[sheet.Cells[1, 30], sheet.Cells[2, 39]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            var size = sheet.Columns[26].ColumnWidth;
            sheet.Columns[26].ColumnWidth = 0;

            size += sheet.Columns[27].ColumnWidth;
            sheet.Columns[27].ColumnWidth = 0;

            size += sheet.Columns[28].ColumnWidth;
            sheet.Columns[28].ColumnWidth = 0;

            size += sheet.Columns[29].ColumnWidth;
            sheet.Columns[29].ColumnWidth = 0;

            sheet.Columns[33].ColumnWidth = sheet.Columns[33].ColumnWidth + size;

            size = 0.6;
            sheet.Columns[33].ColumnWidth = sheet.Columns[33].ColumnWidth - size;
            sheet.Columns[30].ColumnWidth = sheet.Columns[30].ColumnWidth + size;

            sheet.Range[sheet.Cells[1, 30], sheet.Cells[1, 39]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 39]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
            sheet.Range[sheet.Cells[1, 30], sheet.Cells[1, 39]].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 39]].Borders[XlBordersIndex.xlEdgeBottom].Weight;

            sheet.Range[sheet.Cells[1, 30], sheet.Cells[1, 30]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 30]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle;
            sheet.Range[sheet.Cells[1, 30], sheet.Cells[1, 30]].Borders[XlBordersIndex.xlEdgeLeft].Weight = sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 30]].Borders[XlBordersIndex.xlEdgeLeft].Weight;

        }
    }
}
