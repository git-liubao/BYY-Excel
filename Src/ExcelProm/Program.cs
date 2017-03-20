

using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProm
{
    
    class Program
    {
        static void Main(string[] args)
        {
            var str = File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\Setting.json");
            var setting = JsonConvert.DeserializeObject<Setting>(str);
            TaskFactory f = new TaskFactory();
            int index = 1;

            foreach (var id in setting.FileIDs)
            {
                Console.WriteLine();
                Console.WriteLine($"**********************    index: {index++},  id: {id}");
                Task.Run(() => process(setting.FilePath, id, setting.Suffix)).GetAwaiter().GetResult();
            }
        }

        private static void process (string filePath, string fileId, string suffix)
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


            string[] page = Directory.GetFiles(filePath, $"UBS* {fileId}*.xls*");
            if (page.Length != 1)
            {
                Console.WriteLine($"There are {page.Length} pages for {fileId}.");
            }
            if (page.Length == 0)
            {
                return;
            }

            var mergedPage = page[0].Split(".".ToCharArray())[0] + suffix + ".xlsx";
            var ps = new string[] { frontPage[0], page[0] };

            Console.WriteLine($"   front page  :   {frontPage[0]}.");
            Console.WriteLine($"         page  :   {page[0]}");
            Console.WriteLine($"  merged page  :   {mergedPage}");

            MergeWorkbooks(mergedPage, ps);
        }

        private static void MergeWorkbooks(string destinationFilePath, params string[] sourceFilePaths)
        {
            var app = new Application();
            app.DisplayAlerts = false; // No prompt when overriding

            // Create a new workbook (index=1) and open source workbooks (index=2,3,...)
            Workbook destinationWb = app.Workbooks.Add();
            foreach (var sourceFilePath in sourceFilePaths)
            {
                app.Workbooks.Add(sourceFilePath);
            }

            // Copy all worksheets
            Worksheet after = destinationWb.Worksheets[1];
            for (int wbIndex = app.Workbooks.Count; wbIndex >= 2; wbIndex--)
            {
                Workbook wb = app.Workbooks[wbIndex];
                for (int wsIndex = wb.Worksheets.Count; wsIndex >= 1; wsIndex--)
                {
                    Worksheet ws = wb.Worksheets[wsIndex];
                    ws.Copy(After: after);
                }
            }

            // Close source documents before saving destination. Otherwise, save will fail
            for (int wbIndex = 2; wbIndex <= app.Workbooks.Count; wbIndex++)
            {
                Workbook wb = app.Workbooks[wbIndex];
                wb.Close();
            }

            // Delete default worksheet
            after.Delete();

            //format workbook
            FormatWorkBook(destinationWb);

            // Save new workbook
            destinationWb.SaveAs(destinationFilePath);
            destinationWb.Close();

            app.Quit();
        }

        private static void FormatWorkBook(Workbook wb)
        {
            int totalPage = wb.Sheets.Count;

            // resize picture 1's size
            Worksheet sheet1 = wb.Sheets[1];
            sheet1.Cells[3, 24].Value = $"第   1   页    共   {totalPage}   页";
            sheet1.Cells[2, 28].Value = "修改：0";
            Pictures pics = sheet1.Pictures(System.Type.Missing) as Pictures;
            foreach(Picture p in pics)
            {
                if (p.Name == "Picture 1")
                {
                    p.ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;
                    p.ShapeRange.Width = 1.13f;
                    p.ShapeRange.Height = 6.14f;
                    p.ShapeRange.ScaleHeight(1.45f, Microsoft.Office.Core.MsoTriState.msoTrue);
                    p.ShapeRange.ScaleWidth(1.42f, Microsoft.Office.Core.MsoTriState.msoTrue);
                }
            }

            

            for(int i=2;i<=wb.Sheets.Count;i++)
            {
                Worksheet sheet = wb.Sheets[i];

                //reset merge and center for cell AD1
                sheet.Range[sheet.Cells[1, 30], sheet.Cells[2, 30]].UnMerge();

                var size = sheet.Columns[32].ColumnWidth - 0.7;
                sheet.Columns[32].ColumnWidth = 0.7;
                sheet.Columns[33].ColumnWidth = sheet.Columns[33].ColumnWidth + size;

                //Set value to page 2 - end
                sheet1.Cells[1, 24].Copy(sheet.Cells[1, 30]);
                sheet1.Cells[1, 25].Copy(sheet.Cells[1, 33]);
                sheet.Cells[1, 33].Value = "=中文首页!Y1";
                sheet1.Cells[2, 24].Copy(sheet.Cells[2, 30]);
                sheet1.Cells[2, 25].Copy(sheet.Cells[2, 33]);
                sheet.Cells[2, 33].Value = "=中文首页!Y2";
                sheet1.Cells[2, 28].Copy(sheet.Cells[2, 37]);

                sheet.Range[sheet.Cells[3, 30], sheet.Cells[3, 39]].Value = $"第   {i}   页    共   {totalPage}   页";
                sheet.Range[sheet.Cells[3, 30], sheet.Cells[3, 39]].Font.Name = "宋体";
                sheet.Range[sheet.Cells[3, 30], sheet.Cells[3, 39]].Font.Size = 9;

                //sheet.Range[sheet.Cells[1, 30], sheet.Cells[1, 30]].Style.Borders = sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 30]].Style;
                //sheet.Range[sheet.Cells[1, 33], sheet.Cells[1, 33]].Style.Borders = sheet.Range[sheet.Cells[2, 30], sheet.Cells[2, 30]].Style;
                sheet.Cells[1, 33].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Cells[1, 32].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
                sheet.Cells[1, 33].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Cells[1, 32].Borders[XlBordersIndex.xlEdgeBottom].Weight;
                //sheet.Cells[1, 33].Borders[XlBordersIndex.xlEdgeBottom].ColorIndex = sheet.Cells[1, 32].Borders[XlBordersIndex.xlEdgeBottom].ColorIndex;
                sheet.Cells[1, 30].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sheet.Cells[1, 32].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
                sheet.Cells[1, 30].Borders[XlBordersIndex.xlEdgeBottom].Weight = sheet.Cells[1, 32].Borders[XlBordersIndex.xlEdgeBottom].Weight;
                //sheet.Cells[1, 30].Borders[XlBordersIndex.xlEdgeBottom].ColorIndex = sheet.Cells[1, 32].Borders[XlBordersIndex.xlEdgeBottom].ColorIndex;

            }
        }
    }
}
