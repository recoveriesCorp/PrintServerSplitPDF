//   PrintServerSplitPDF is a command line utility to split multi-document pdf files into single document pdf files.
//   Copyright(C) 2018 recoveriesCorp
//
//   This program is free software: you can redistribute it and/or modify
//   it under the terms of the GNU Affero General Public License as
//   published by the Free Software Foundation, either version 3 of the
//   License, or(at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//   but WITHOUT ANY WARRANTY; without even the implied warranty of
//   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
//   GNU Affero General Public License for more details.
//
//   You should have received a copy of the GNU Affero General Public License
//   along with this program.If not, see<http://www.gnu.org/licenses/>.

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace PrintServerSplitPDF
{
    class Helper
    {
        public static Logger _logger = LogManager.GetLogger("other");

        public static void LoggerConfig(string logDir, string logFilename)
        {
            if (!Directory.Exists(logDir))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(logDir);
                }
                catch
                {
                    LogMsg(null, string.Format("Can not create dir - please check folder access: {0}", new object[] { logDir }));
                }
            }

            // Step 1. Create configuration object 
            var config = new LoggingConfiguration();

            // Step 2. Create targets and add them to the configuration 
            var consoleTarget = new ColoredConsoleTarget();
            config.AddTarget("console", consoleTarget);

            var fileTarget = new FileTarget();
            config.AddTarget("file", fileTarget);

            // Step 3. Set target properties 
            consoleTarget.Layout = @"${date:format=HH\:mm\:ss} ${logger} ${message}";
            fileTarget.FileName = System.IO.Path.Combine(logDir, logFilename);
            //fileTarget.Layout = "${message}";

            // Step 4. Define rules
            var rule1 = new LoggingRule("*", LogLevel.Debug, consoleTarget);
            config.LoggingRules.Add(rule1);

            var rule2 = new LoggingRule("*", LogLevel.Debug, fileTarget);
            config.LoggingRules.Add(rule2);

            // Step 5. Activate the configuration
            LogManager.Configuration = config;
            //LogManager.ReconfigExistingLoggers();
        }

        public static void LogMsg(string logFilePath, string msg = null, string type = null)
        {
            try
            {
                string errLog = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SplitPDFLogs");
                if (!Directory.Exists(errLog))
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(errLog);
                    }
                    catch
                    {

                    }
                }

                if (logFilePath==null)
                {
                    int maxWait = 1;
                    DateTime dtStart = DateTime.Now;
                    string fileName = dtStart.ToString("yyyy-MM-dd") + ".log";

                    string lockFile = System.IO.Path.Combine(errLog, "lock.file");
                    while (true)
                    {
                        string logFile = System.IO.Path.Combine(errLog, fileName);
                        if (!File.Exists(lockFile))
                        {
                            File.Create(lockFile).Close();
                            if (type == "Info")
                            {
                                _logger.Info(msg);
                            }
                            else
                            {
                                _logger.Error(msg);
                            }
                            File.Delete(lockFile);
                            break;
                        }
                        else
                        {
                            Thread.Sleep(500);
                            if (DateTime.Now.Subtract(dtStart).Minutes > maxWait)
                            {
                                Console.WriteLine(string.Format("LogMsg is taking too long, over {0} minutes. Check if lock file exists at {1}", maxWait, lockFile));
                                throw new Exception(string.Format("LogMsg is taking too long, over {0} minutes. Check if lock file exists at {1}", maxWait, lockFile));
                            }
                        }
                    }
                }
                else
                {
                    try
                    {
                        File.AppendAllText(logFilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + "   " + Thread.CurrentThread.ManagedThreadId + "|INFO " + msg + Environment.NewLine);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch
            {
            }
        }

        public static bool SplitPDFs(string pdfFilePath, string tmpControlFileName, string docType, string database, string logFilePath)
        {
            bool retVal = false;
            try
            {
                string searthText = "CODENO[";
                LogMsg(logFilePath, string.Format("SplitPDFs Starting", new object[] { pdfFilePath }));
                List<int> intervals = ReadPdfFile(pdfFilePath, searthText, logFilePath);
                FileInfo finfo = new FileInfo(pdfFilePath);
                List<int> queueList = new List<int>();
                LogMsg(logFilePath, string.Format("SplitPDFs Marker 02", new object[] { }));
                string redmapFolder = System.IO.Path.Combine(finfo.Directory.Parent.FullName, "Redmap");

                string tmpPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp", "Redmap");

                string docCode2 = null;
                string jobNo = null;
                string requestType = null;
                try
                {
                    LogMsg(logFilePath, string.Format("SplitPDFs Marker 03", new object[] { }));
                    string[] nameParts = finfo.Name.Split('-');
                    docCode2 = nameParts[0];
                    jobNo = nameParts[1];
                    requestType = nameParts[0].Split('_')[1];
                    LogMsg(logFilePath, string.Format("SplitPDFs Marker 04", new object[] { }));
                }
                catch (Exception ex1)
                {
                    LogMsg(logFilePath, string.Format("SplitPDFs nameParts split failed : Error Msg: {0} Trace: {1}", new object[] { ex1.Message, ex1.StackTrace }));
                    throw new Exception();
                }
                LogMsg(logFilePath, string.Format("SplitPDFs Marker 06", new object[] { }));
                int pageNameSuffix = 0;
                int pageNumber = 1;
                FileInfo file = new FileInfo(pdfFilePath);
                LogMsg(logFilePath, string.Format("SplitPDFs Marker 07", new object[] { }));

                List<string> linesControl = null;
                string sep = "\t";
                string dbCode = null;
                string dhCode = null;
                string seqNo = null;
                string strDatetime = DateTime.Now.ToString("dd_MM_yy-HH_mm_ss_ffffff");
                string fileName = System.IO.Path.GetFileNameWithoutExtension(pdfFilePath);

                LogMsg(logFilePath, string.Format("SplitPDFs before starting split loop of: {0} iterations", new object[] { intervals.Count }));
                for (int i = 0; i < intervals.Count; i++)
                {
                    try
                    {
                        //LogMsg(logFilePath, string.Format("SplitPDFs Iteration {0}", new object[] { i + 1 }));
                        linesControl = File.ReadLines(tmpControlFileName, Encoding.GetEncoding(1252)).Skip(i).Take(1).ToList();
                        try
                        {

                            string[] splitControlFile = linesControl[0].Split(sep.ToCharArray());
                            dbCode = splitControlFile[2];
                            seqNo = splitControlFile[11].Split('-')[0];
                            dhCode = splitControlFile[12]; // should be this one                                                
                        }
                        catch (Exception ex2)
                        {
                            LogMsg(logFilePath, string.Format("SplitPDFs splitControlFile split failed : Error Msg: {0} Trace: {1}", new object[] { ex2.Message, ex2.StackTrace }));
                            throw new Exception();
                        }
                        string newPdfFileName = string.Format("{0}#{1}#ACT0#ACC={2}#{3}-{4}-{5}-{6}-{7}", dbCode, docType, database, docCode2, dhCode, seqNo, jobNo, strDatetime);
                        pageNameSuffix++;
                        int size = 0;

                        // Intialize a new PdfReader instance with the contents of the source Pdf file:  
                        using (PdfReader reader = new PdfReader(pdfFilePath))
                        {
                            // special size formula for last record in file
                            if (i == intervals.Count - 1)
                            {
                                size = (reader.NumberOfPages + 1) - intervals[i];
                            }

                            // size formula for records in file
                            else
                            {
                                size = intervals[i + 1] - intervals[i];
                            }
                        }

                        LogMsg(logFilePath, string.Format("SplitPDFs starting SplitAndSaveInterval iteration {0}", new object[] { i + 1 }));
                        SplitAndSaveInterval(pdfFilePath, redmapFolder, pageNumber, size, newPdfFileName, fileName, logFilePath, queueList, requestType);
                        LogMsg(logFilePath, string.Format("SplitPDFs finished  SplitAndSaveInterval iteration {0}", new object[] { i + 1 }));

                        pageNumber += size;
                        //}
                    }
                    catch (Exception ex3)
                    {
                        LogMsg(logFilePath, string.Format("SplitPDFs split failed : Error Msg: {0} Trace: {1}", new object[] { ex3.Message, ex3.StackTrace }));
                        throw new Exception();
                    }
                    LogMsg(logFilePath, string.Format("SplitPDFs completed split loop", new object[] { }));
                }
                retVal = true;
            }
            catch (Exception ex)
            {
                LogMsg(logFilePath, string.Format("SplitPDFs: Error Msg: {0} Trace: {1}", new object[] { ex.Message, ex.StackTrace }));
            }
            LogMsg(logFilePath, string.Format("SplitPDFs before return", new object[] { pdfFilePath }));

            return retVal;
        }

        // SplitAndSaveInterval method  saves number of given pages as a single pdf document (i.e. creates single letter)
        // For "B" jobs it also re-arranges letters based on their number of pages and merges them back to single file, which is later sent to printer
        private static void SplitAndSaveInterval(string pdfFilePath, string outputPath, int startPage, int interval, string pdfFileName, string fileName, string logFilePath, List<int> queueList, string requestType)
        {
            try
            {
                //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 01", new object[] { pdfFilePath }));
                FileInfo finfo = new FileInfo(pdfFilePath);
                // number large enough to count different lengths of letters. Atm 3 would be enough i.e. 1,2 and 3 page letters.
                PdfCopy[] PdfCopyArray = new PdfCopy[10];
                string assistFolder = System.IO.Path.Combine(finfo.Directory.Parent.FullName, "AssistBatch");
                string batchQueueFolder = System.IO.Path.Combine(finfo.Directory.Parent.FullName, "BatchQueues");

                //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 02", new object[] { }));
                bool first = false;
                bool addBlank = false;
                int queue = 1;
                List<int> neededPages = new List<int>();
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 03", new object[] { }));
                    Document document = new Document();

                    //prep work for batch print queues, to merge letters of same length
                    if (requestType == "B")
                    {
                        LogMsg(logFilePath, string.Format("SplitAndSaveInterval start splitting B type letter", new object[] { }));
                        if (interval % 2 == 0) queue = (interval / 2);
                        else
                        {
                            //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 05", new object[] { }));
                            queue = ((interval + 1) / 2);
                            addBlank = true;
                        }
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 06", new object[] { }));
                        fileName = queue + "^" + fileName;
                        if (!queueList.Contains(queue))
                        {
                            first = true;
                            queueList.Add(queue);
                            PdfCopyArray[queue] = new PdfCopy(document, new FileStream(batchQueueFolder + "\\" + fileName + ".pdf", FileMode.Create));
                        }
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 07", new object[] { }));
                        PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 08", new object[] { }));
                        PdfCopy assistCopy = new PdfCopy(document, new FileStream(assistFolder + "\\" + pdfFileName + ".pdf", FileMode.Create));
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 09", new object[] { }));
                        document.Open();
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 10", new object[] { }));
                        for (int pagenumber = startPage; pagenumber < (startPage + interval); pagenumber++)
                        {
                            neededPages.Add(pagenumber);
                            if (reader.NumberOfPages >= pagenumber)
                            {
                                copy.AddPage(copy.GetImportedPage(reader, pagenumber));

                                assistCopy.AddPage(copy.GetImportedPage(reader, pagenumber));

                                if (first == true && requestType == "B")
                                {
                                    PdfCopyArray[queue].AddPage(copy.GetImportedPage(reader, pagenumber));
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 11", new object[] { }));
                        if (addBlank == true)
                        {
                            //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 12", new object[] { }));
                            assistCopy.AddPage(PageSize.A4, 0);
                        }
                        if (first == true && requestType == "B" && addBlank == true)
                        {
                            //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 13", new object[] { }));
                            PdfCopyArray[queue].AddPage(PageSize.A4, 0);
                        }
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 14", new object[] { }));
                        document.Close();
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 15", new object[] { }));
                        reader.Close();
                        LogMsg(logFilePath, string.Format("SplitAndSaveInterval finished splitting B type letter", new object[] { }));
                    }
                    else
                    {
                        LogMsg(logFilePath, string.Format("SplitAndSaveInterval start splitting non-B type letter", new object[] { }));
                        PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 18", new object[] { }));
                        document.Open();
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 19", new object[] { }));
                        for (int pagenumber = startPage; pagenumber < (startPage + interval); pagenumber++)
                        {
                            neededPages.Add(pagenumber);
                            if (reader.NumberOfPages >= pagenumber)
                            {
                                copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                            }

                            else
                            {
                                break;
                            }
                        }
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 20", new object[] { }));
                        document.Close();
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 21", new object[] { }));
                        reader.Close();
                        LogMsg(logFilePath, string.Format("SplitAndSaveInterval ended splitting non-B type letter", new object[] { }));
                    }
                    //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 23", new object[] { }));
                }
                //LogMsg(logFilePath, string.Format("SplitAndSaveInterval closing PdfReader reader", new object[] { }));
                if (first == false && requestType == "B")
                {
                    LogMsg(logFilePath, string.Format("SplitAndSaveInterval start extra processing for 'B' type job", new object[] { }));
                    string queuePDF = System.IO.Path.Combine(batchQueueFolder, fileName + ".pdf");
                    string tmpPDF = System.IO.Path.Combine(batchQueueFolder, fileName + "_tmp.pdf");
                    string letter = System.IO.Path.Combine(assistFolder, pdfFileName + ".pdf");

                    if (!File.Exists(tmpPDF))
                    {
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 26", new object[] { }));
                        File.Create(tmpPDF).Close();
                    }
                    //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 27", new object[] { }));
                    File.Copy(queuePDF, tmpPDF, true);
                    //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 28", new object[] { }));
                    List<string> sourceFiles = new List<string>();
                    sourceFiles.Add(tmpPDF);
                    sourceFiles.Add(letter);
                    LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 29", new object[] { }));
                    if (!MergePDFs(sourceFiles, queuePDF, logFilePath)) throw new Exception();
                    LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 30", new object[] { }));
                    if (File.Exists(tmpPDF))
                    {
                        //LogMsg(logFilePath, string.Format("SplitAndSaveInterval Marker 31", new object[] { }));
                        File.Delete(tmpPDF);
                    }
                    //LogMsg(logFilePath, string.Format("SplitAndSaveInterval finished extra processing for 'B' type job", new object[] { }));
                }
                //LogMsg(logFilePath, string.Format("SplitAndSaveInterval before returning", new object[] { }));
            }
            catch (Exception ex)
            {
                LogMsg(logFilePath, string.Format("SplitAndSaveInterval: Error Msg: {0} Trace: {1}", new object[] { ex.Message, ex.StackTrace }));

                // clean up temp redmap folder in case of error, to fail completion check in main PrintServer
                try
                {
                    foreach (FileInfo file in new DirectoryInfo(outputPath).GetFiles())
                    {
                        file.Delete();
                    }
                }
                catch
                {

                }
                throw new Exception();
            }
        }

        public static List<int> ReadPdfFile(string pdfFilePath, String searthText, string logFilePath)
        {
            List<int> pages = new List<int>();
            try
            {
                LogMsg(logFilePath, string.Format("ReadPdfFile Marker 01", new object[] { pdfFilePath }));
                FileInfo finfo = new FileInfo(pdfFilePath);

                LogMsg(logFilePath, string.Format("ReadPdfFile Marker 02", new object[] { }));
                if (File.Exists(pdfFilePath))
                {
                    LogMsg(logFilePath, string.Format("ReadPdfFile Marker 03", new object[] { }));
                    PdfReader pdfReader = new PdfReader(pdfFilePath);
                    for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                        string currentPageText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                        if (currentPageText.Contains(searthText))
                        {
                            pages.Add(page);
                        }
                    }
                    pdfReader.Close();
                    LogMsg(logFilePath, string.Format("ReadPdfFile completed full method", new object[] { }));
                }
                LogMsg(logFilePath, string.Format("ReadPdfFile before returning", new object[] { }));
                return pages;
            }
            catch (Exception ex)
            {
                LogMsg(logFilePath, string.Format("ReadPdfFile of SplitPDF: Error Msg: {0} Trace: {1}", new object[] { ex.Message, ex.StackTrace }));
                return pages;
            }
        }

        public static bool MergePDFs(List<string> sourceFiles, string targetFile, string logFilePath)
        {
            bool retVal = false;
            try
            {
                LogMsg(logFilePath, string.Format("MergePDFs Marker 01", new object[] { }));
                FileInfo finfo = new FileInfo(targetFile);
                LogMsg(logFilePath, string.Format("MergePDFs Marker 02", new object[] { }));
                using (FileStream stream = new FileStream(targetFile, FileMode.Create))
                {
                    LogMsg(logFilePath, string.Format("MergePDFs Marker 03", new object[] { }));
                    Document document = new Document();
                    //PdfCopy pdf = new PdfCopy(document, stream);
                    PdfSmartCopy pdf = new PdfSmartCopy(document, stream);
                    LogMsg(logFilePath, string.Format("MergePDFs Marker 04", new object[] { }));
                    //pdf.SetPdfVersion(PdfWriter.PDF_VERSION_1_3);
                    //pdf.PDFXConformance = PdfWriter.PDFX1A2001; 
                    using (PdfReader reader = null)
                    {
                        try
                        {
                            LogMsg(logFilePath, string.Format("MergePDFs Marker 05", new object[] { }));
                            document.Open();
                            foreach (string file in sourceFiles)
                            {
                                //reader = new PdfReader(file);
                                //pdf.AddDocument(reader);
                                //reader.Close();
                                //Bring in every page from the old PDF
                                using (var r = new PdfReader(file))
                                {
                                    for (var i = 1; i <= r.NumberOfPages; i++)
                                    {
                                        pdf.AddPage(pdf.GetImportedPage(r, i));
                                    }
                                }
                            }
                            LogMsg(logFilePath, string.Format("MergePDFs completed adding all pages", new object[] { }));
                            retVal = true;
                        }
                        catch (Exception ex)
                        {
                            LogMsg(logFilePath, string.Format("MergePDFs Marker 07 - Exception", new object[] { }));
                            retVal = false;
                            if (reader != null)
                            {
                                reader.Close();
                            }
                            LogMsg(logFilePath, string.Format("SplitPDFs: Error Msg: {0} Trace: {1}", new object[] { ex.Message, ex.StackTrace }));
                            throw new Exception();
                        }
                        finally
                        {
                            LogMsg(logFilePath, string.Format("MergePDFs closing document", new object[] { }));
                            if (document != null)
                            {
                                document.Close();
                            }
                        }
                    }
                    LogMsg(logFilePath, string.Format("MergePDFs closing PdfReader reader", new object[] { }));
                }
                LogMsg(logFilePath, string.Format("MergePDFs before returning", new object[] { }));
                return retVal;
            }
            catch (Exception ex)
            {
                LogMsg(logFilePath, string.Format("MergePDFs of SplitPDF: Error Msg: {0} Trace: {1}", new object[] { ex.Message, ex.StackTrace }));
                return false;
            }
        }
    }
}