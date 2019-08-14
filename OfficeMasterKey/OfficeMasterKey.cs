//===================================================================================================================//
//-------------------------------------------------------------------------------------------------------------------//
//                                                                                                                   //
//  OfficeMasterKey                                                                                                  //
//                                                                                                                   //
//  Remove document protection mechanisms for DOCX and XLSX files.                                                   //
//                                                                                                                   //
//  https://github.com/ProcessBolt/OfficeMasterKey                                                                   //
//                                                                                                                   //
//-------------------------------------------------------------------------------------------------------------------//
//                                                                                                                   //
//  Copyright 2019 ProcessBolt, Inc.                                                                                 //
//                                                                                                                   //
//  Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated     //
//  documentation files (the “Software”), to deal in the Software without restriction, including without limitation  //
//  the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,     //
//  and to permit persons to whom the Software is furnished to do so, subject to the following conditions:           //
//                                                                                                                   //
//  The above copyright notice and this permission notice shall be included in all copies or                         //
//  substantial portions of the Software.                                                                            //
//                                                                                                                   //
//  THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED    //
//  TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL     //
//  THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF    //
//  CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER         //
//  DEALINGS IN THE SOFTWARE                                                                                         //
//                                                                                                                   //
//-------------------------------------------------------------------------------------------------------------------//
//
//  HISTORY:
//      2019-08-12 Original (Dan Gardner, ProcessBolt)
//      2019-08-14 Split CLI and OfficeMasterKey library (Dan Gardner, ProcessBolt)
//
//

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeMasterKey
{
    public class MasterKey
    {
        /// <summary>
        /// Remove the document protection elements from an XLSX (SpreadsheetDocument) file
        /// </summary>
        /// <param name="filename">XLSX file to remove protection</param>
        public void UnprotectXlsx(string filename)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, true))
            {
                spreadsheetDocument.WorkbookPart.Workbook.RemoveAllChildren<WorkbookProtection>();

                foreach (WorksheetPart worksheetPart in spreadsheetDocument.WorkbookPart.WorksheetParts)
                {
                    worksheetPart.Worksheet.RemoveAllChildren<SheetProtection>();
                }

                spreadsheetDocument.Save();

                spreadsheetDocument.Close();
            }
        }

        /// <summary>
        /// Remove the document protection elements from an DOCX (WordprocessingDocument) file
        /// </summary>
        /// <param name="filename">DOCX file to remove protection</param>
        public void UnprotectDocx(string filename)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filename, true))
            {
                wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren<DocumentProtection>();

                wordprocessingDocument.Save();

                wordprocessingDocument.Close();
            }
        }

        /// <summary>
        /// Determines if a file is a valid XLSX (OpenXML SpreadsheetDocument) document.
        /// </summary>
        /// <param name="filename">Name of file to query</param>
        /// <returns>True if valid XLSX file.</returns>
        public bool FileIsXlsx(string filename)
        {
            try
            {
                /* Try to open file as a SpreadsheetDocument */
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, true))
                {
                    return true;
                }
            }
            catch (System.IO.FileFormatException)
            {
                /* Is not an OpenXML container */
                return false;
            }
            catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
            {
                /* It is an OpenXML container, but does not have the correct parts */
                return false;
            }
        }

        /// <summary>
        /// Determines if a file is a valid DOCX (OpenXML WordprocessingDocument) document.
        /// </summary>
        /// <param name="filename">Name of file to query</param>
        /// <returns>True if valid DOCX file.</returns>
        public bool FileIsDocx(string filename)
        {
            try
            {
                /* Try to open file as a WordprocessingDocument */
                using (WordprocessingDocument spreadsheetDocument = WordprocessingDocument.Open(filename, true))
                {
                    return true;
                }
            }
            catch (System.IO.FileFormatException)
            {
                /* Is not an OpenXML container */
                return false;
            }
            catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
            {
                /* It is an OpenXML container, but does not have the correct parts */
                return false;
            }
        }

    }
}
