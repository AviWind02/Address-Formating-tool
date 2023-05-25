/*
 * Create By Avi 
 * 2023 - 05 -25
 * 

MIT License

Copyright (c) 2023 Aviraj Gill

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE. 

 */
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Xml.Linq;


namespace Address_Format.Class
{
    internal class creatingwordfile
    {

        public void CreateWordDocument(List<CompanyInfo> companyList, string filePath)
        {
            // Create a new Word application and document
            Application wordApp = new Application();
            Document wordDoc = wordApp.Documents.Add();

            // Set the page size and margins for portrait orientation
            wordDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
            wordDoc.PageSetup.PageWidth = wordApp.MillimetersToPoints(210f);   // A4 width in mm
            wordDoc.PageSetup.PageHeight = wordApp.MillimetersToPoints(297f);  // A4 height in mm
            // Set the margins
            wordDoc.PageSetup.TopMargin = wordApp.MillimetersToPoints(15.1f);   // 1.51cm
            wordDoc.PageSetup.BottomMargin = wordApp.MillimetersToPoints(13f);   // 1.3cm
            wordDoc.PageSetup.LeftMargin = wordApp.MillimetersToPoints(8.6f);    // 0.86cm
            wordDoc.PageSetup.RightMargin = wordApp.MillimetersToPoints(7.9f);   // 0.79cm
            wordDoc.PageSetup.Gutter = 0;
            wordDoc.PageSetup.GutterPos = WdGutterStyle.wdGutterPosLeft;



            // Calculate the number of rows and columns based on the companyList size
            int rowCount = (int)Math.Ceiling((double)companyList.Count / 3);
            int columnCount = 3;

            // Add a table to the document
            Table table = wordDoc.Tables.Add(wordDoc.Range(), rowCount, columnCount);

            // Set the cell dimensions and formatting
            table.AllowAutoFit = false;
            for (int i = 1; i <= table.Columns.Count; i++)
            {
                table.Columns[i].Width = wordApp.MillimetersToPoints(63.5f); // 63.5mm
            }
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                table.Rows[i].Height = wordApp.MillimetersToPoints(38.1f); // 38.1mm
            }

            // Set the table style to "Table Grid"
            table.set_Style("Table Grid");

            // Set the table borders color to white
            table.Borders.OutsideColor = WdColor.wdColorWhite;
            table.Borders.InsideColor = WdColor.wdColorWhite;



            // Populate the table with company information
            int companyIndex = 0;
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    if (companyIndex < companyList.Count)
                    {
                        CompanyInfo companyInfo = companyList[companyIndex];

                        // Create a new cell range and apply formatting
                        Range cellRange = table.Cell(row, column).Range;

                        // Set horizontal alignment to center
                        cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        // Insert the company information into the cell
                        cellRange.Text = $"{companyInfo.CompanyName}\n" +
                                         $"{companyInfo.AddressLine1}\n" +
                                         $"{companyInfo.AddressLine2}\n" +
                                         $"{companyInfo.County}\n" +
                                         $"{companyInfo.PostalCode}";

                        companyIndex++;
                    }
                }
            }

            // Save the document
            wordDoc.SaveAs2(filePath);

            // Close the document and Word application
            wordDoc.Close();
            wordApp.Quit();
        }
    }
}



