//===================================================================================================================//
//-------------------------------------------------------------------------------------------------------------------//
//                                                                                                                   //
//  OfficeMasterKey CLI                                                                                              //
//                                                                                                                   //
//  Remove document protection mechanisms for DOCX and XLSX files.                                                   //
//                                                                                                                   //
//  https://github.com/ProcessBolt/OfficeMasterKey                                                                   //
//                                                                                                                   //
//-------------------------------------------------------------------------------------------------------------------//
//                                                                                                                   //
//  Copyright 2019-2020 ProcessBolt, Inc.                                                                            //
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
//      2019-08-15 Added XxxxIsProtected methods (Dan Gardner, ProcessBolt)
//
//

using System;
using OfficeMasterKey;

namespace omkcli
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("");
            Console.WriteLine("Office Master Key");
            Console.WriteLine("Remove document protection mechanisms for DOCX and XLSX files.");
            Console.WriteLine("");

            if (args.Length == 0)
            {
                Console.WriteLine("Usage:  omkcli file1 [file2 ... [ fileN ] ] ");
                Console.WriteLine("");

                return;
            }

            MasterKey masterKey = new MasterKey();

            foreach (string arg in args)
            {
                Console.Write("Removing protection from: " + arg);

                try
                {
                    if (masterKey.FileIsXlsx(arg))
                    {
                        if (masterKey.XlsxIsProtected(arg))
                        {
                            string protectedParts = "";

                            if (masterKey.XlsxIsWorkbookProtected(arg) && masterKey.XlsxIsWorksheetProtected(arg))
                            {
                                protectedParts = "[Workbook,Worksheet]";
                            }
                            else if (masterKey.XlsxIsWorkbookProtected(arg))
                            {
                                protectedParts = "[Workbook]";
                            }
                            else if (masterKey.XlsxIsWorksheetProtected(arg))
                            {
                                protectedParts = "[Worksheet]";
                            }

                            masterKey.UnprotectXlsx(arg);
                            Console.WriteLine("  OK. " + protectedParts);
                        }
                        else
                        {
                            Console.WriteLine("  Not protected.");
                        }
                    }
                    else if (masterKey.FileIsDocx(arg))
                    {
                        if (masterKey.DocxIsProtected(arg))
                        {
                            masterKey.UnprotectDocx(arg);
                            Console.WriteLine("  OK.");
                        }
                        else
                        {
                            Console.WriteLine("  Not protected.");
                        }
                    }
                    else
                    {
                        throw new ApplicationException("Not recognized as valid DOCX or XLSX file type.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  Failed.  " + ex.Message);
                }
            }
        }

    }

}

//-------------------------------------------------------------------------------------------------------------------//
// EOF                                                                                                               //
//-------------------------------------------------------------------------------------------------------------------//
