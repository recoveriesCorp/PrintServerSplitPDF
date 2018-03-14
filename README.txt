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
//   along with this program. If not, see<http://www.gnu.org/licenses/>.

****************************
PrintServer SplitPDF 
****************************
PrintServer SplitPDF  is a utility to split a large multi-document PDF file into many single-document PDF files by searching for the "CODENO[" string as a marker, separating documents.
This program is utilising:
- iText library, distributed under AGPL (https://itextpdf.com/AGPL);
- NLOG library (Copyright (c) 2004-2018 Jaroslaw Kowalski <jaak@jkowalski.net>, Kim Christensen, Julian Verdurmen), distributed under BSD 3-Clause "New" or "Revised" License.
****************************
The program is a C# console app, which is designed to work with 5 command line arguments:

pdfFileInPath - full path of the pdf file to be split
supportFilePath - full path of the supplementary tab delimited file, which contains additional info about each letter (used to create meaningful and unique filename for each letter)
docType - becomes part of filename of new files
database - becomes part of filename of new files
logFilePath - full path of the external log file to add logs. If the provided logfile does not exist, will log in the installation directory.

C# Use example:
****************************
string pdfFileInPath = @"C:\Temp\Printing\XYZDOC_D-999999-20180103135603-00001_00008_00008.pdf";
string supportFilePath = @"C:\Temp\PrintSupport\XYZDOC_D-999999-20180103135603-00001_00008_00008.DAT";
string docType = "LETTER";
string database = "DB";
string logFilePath = @"C:\Temp\SplitPDFLogs\XYZDOC_D-999999-20180103135603-00001_00008_00008.log";
string parameteres = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\" \"{4}\"", pdfFileInPath, supportFilePath, docType, database, logFilePath);
p = Process.Start(exePath, parameteres);
****************************