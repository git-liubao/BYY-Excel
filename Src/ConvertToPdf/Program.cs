using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;
using Newtonsoft.Json;

namespace ConvertToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            var str = File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\Setting.json");
            var paths = JsonConvert.DeserializeObject<String[]>(str);
            foreach(var path in paths)
            {
                Process(path, 0);
            }
        }

        public static void Process(string path, int level)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            FileInfo fil = new FileInfo(path);

            if (fil == null)
            {
                return;
            }


            if (fil.Attributes == FileAttributes.Directory)//判断类型是选择的是文件夹还是文件
            {
                PrintStar(level, path);

                string[] folders = Directory.GetDirectories(path);
                foreach (var folder in folders)
                {
                    Process(folder, level +1);
                }

                string[] files = Directory.GetFiles(path);
                foreach(var file in files)
                {
                    Process(file, level + 1);
                }
            }
            else if(fil.Attributes == FileAttributes.Archive)
            {
                switch(fil.Extension)
                {
                    case ".doc":
                    case ".dcox":
                        PrintStar(level, path);

                        DOCConvertToPDF(fil.FullName, GeneratePdfName(fil.FullName));
                        break;
                    case ".xls":
                    case ".xlsx":
                        PrintStar(level, path);

                        XLSConvertToPDF(fil.FullName, GeneratePdfName(fil.FullName));
                        break;
                    case ".ppt":
                    case ".pptx":
                        PrintStar(level, path);

                        PPTConvertToPDF(fil.FullName, GeneratePdfName(fil.FullName));
                        break;
                }
            }
        }

        private static void PrintStar(int num, string name)
        {
            for(int i = 0;i<num;i++)
            {
                Console.Write("--");
            }
            Console.WriteLine((new FileInfo(name)).Name);
        }

        private static string GeneratePdfName(string fullName)
        {
            return fullName.Split(".".ToCharArray())[0] + ".pdf";
        }

        //Word转换成pdf  
        /// <summary>  
        /// 把Word文件转换成为PDF格式文件  
        /// </summary>  
        /// <param name="sourcePath">源文件路径</param>  
        /// <param name="targetPath">目标文件路径</param>   
        /// <returns>true=转换成功</returns>  
        private static bool DOCConvertToPDF(string sourcePath, string targetPath)
        {
            File.Delete(targetPath);

            bool result = false;
            Word.WdExportFormat exportFormat = Word.WdExportFormat.wdExportFormatPDF;
            object paramMissing = Type.Missing;
            Word.Application wordApplication = new Word.Application();
            Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                Word.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                Word.WdExportOptimizeFor paramExportOptimizeFor = Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                Word.WdExportCreateBookmarks paramCreateBookmarks = Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                wordDocument = wordApplication.Documents.Open(
                ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing, ref paramMissing, ref paramMissing,
                ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                    paramExportFormat, paramOpenAfterExport,
                    paramExportOptimizeFor, paramExportRange, paramStartPage,
                    paramEndPage, paramExportItem, paramIncludeDocProps,
                    paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                    paramBitmapMissingFonts, paramUseISO19005_1,
                    ref paramMissing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        /// <summary>  
        /// 把Excel文件转换成PDF格式文件  
        /// </summary>  
        /// <param name="sourcePath">源文件路径</param>  
        /// <param name="targetPath">目标文件路径</param>   
        /// <returns>true=转换成功</returns>  
        private static bool XLSConvertToPDF(string sourcePath, string targetPath)
        {
            File.Delete(targetPath);

            bool result = false;
            Excel.XlFixedFormatType targetType = Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;
            Excel.Application application = null;
            Excel.Workbook workBook = null;
            try
            {
                application = new Excel.Application();
                object target = targetPath;
                object type = targetType;
                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);

                workBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        /// <summary>  
        /// 把PowerPoing文件转换成PDF格式文件  
        /// </summary>  
        /// <param name="sourcePath">源文件路径</param>  
        /// <param name="targetPath">目标文件路径</param>   
        /// <returns>true=转换成功</returns>  
        private static bool PPTConvertToPDF(string sourcePath, string targetPath)
        {
            File.Delete(targetPath);

            bool result;
            PowerPoint.PpSaveAsFileType targetFileType = PowerPoint.PpSaveAsFileType.ppSaveAsPDF;
            object missing = Type.Missing;
            PowerPoint.Application application = null;
            PowerPoint.Presentation persentation = null;
            try
            {
                application = new PowerPoint.Application();
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);

                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
    }
}
