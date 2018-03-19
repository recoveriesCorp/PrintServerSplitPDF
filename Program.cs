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

using NLog;
using NLog.Targets;
using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace PrintServerSplitPDF
{
    class Program
    {
        static int Main(string[] args)
        {
            string pdfFileInPath = null;
            string supportFilePath = null;
            string docType = null;
            string database = null;
            string logFilePath = null;
            bool isSplit = false;
            Console.WriteLine(string.Format("Processing following {0} args: {1}{2}", new object[] { args.Count(), Environment.NewLine, string.Join(Environment.NewLine, args) }));

            switch (args.Count())
            {
                case 0: // test case
                    pdfFileInPath = @"";
                    supportFilePath = @"";
                    docType = "LETTER";
                    database = "";
                    Helper.LogMsg(null, "Zero Args - test case");
                    if (File.Exists(pdfFileInPath) && File.Exists(supportFilePath))
                    {
                        FileInfo finfo1 = new FileInfo(pdfFileInPath);
                        string errLog1 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SplitPDFLogs");
                        string timestamp1 = DateTime.UtcNow.ToString("yyyy-MM-dd_HH-mm-ss-fff_", CultureInfo.InvariantCulture);
                        string fileName1 = timestamp1 + finfo1.Name + ".log";
                        string logFilePath1 = System.IO.Path.Combine(errLog1, fileName1);
                        string line = Environment.NewLine;
                        string msgg = "Received following 5 command string args: " + line;
                        msgg += pdfFileInPath + line + supportFilePath + line + docType + line + database + line + logFilePath1 + line;
                        Helper.LogMsg(null, msgg);
                        //Helper.LogInstanceMsg(instanceLog, msgg);
                        isSplit = Helper.SplitPDFs(null, supportFilePath, docType, database, logFilePath1);
                    }
                    break;
                case 5: // expected use case
                    pdfFileInPath = args[0];
                    supportFilePath = args[1];
                    docType = args[2];
                    database = args[3];
                    logFilePath = args[4];

                    string msg = string.Format("Processing following {0} command string args: {1}{2}", new object[] { args.Count(), Environment.NewLine, string.Join(Environment.NewLine, args) });
                    Helper.LogMsg(logFilePath, msg, "Info");
                    isSplit = Helper.SplitPDFs(pdfFileInPath, supportFilePath, docType, database, logFilePath);
                    Helper.LogMsg(logFilePath, string.Format("Finished SplitPDF for file: {0}", new object[] { pdfFileInPath }));
                    break;

                default:
                    Helper.LogMsg(null, string.Format("Must provide 5 command string args: pdfFileInPath, supportFilePath, docType, database and logFilePath. Provided: {0}. Args: {1}", new object[] { args.Count(), string.Join("|#NEXT#|", args) }));
                    break;
            }
            return isSplit == true ? 0 : 1;
        }
    }
}