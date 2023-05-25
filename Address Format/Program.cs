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
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Address_Format.Class;
using Microsoft.Office.Interop.Word;

namespace Address_Format
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Specify the CSV file path
            Console.WriteLine("Enter a file path:");
            string filePath = Console.ReadLine();
            //string filePath = @"D:\Test.CSV";//Testing

            bool isFileValid = false;
            List<CompanyInfo> companyList = new List<CompanyInfo>();
            creatingwordfile creatingwordfile = new creatingwordfile();

            do
            {
                // Check if the file path is valid and the file extension is ".csv"
                if (File.Exists(filePath) && Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    isFileValid = true;
                    Console.WriteLine("Valid CSV file path entered: " + filePath);

                    try
                    {
                        // Perform further operations on the CSV file.
                        // For example, you can read the contents of the file using StreamReader.
                        using (StreamReader reader = new StreamReader(filePath))
                        {
                            int lineNumber = 1;
                            string line;

                            while ((line = reader.ReadLine()) != null)
                            {
                                Console.WriteLine($"Line Number: {lineNumber}: {line}");
                                lineNumber++;

                                // Split the line into parts using comma as the delimiter
                                string[] parts = line.Split(',');

                                // Create a new CompanyInfo object and assign values to its properties
                                CompanyInfo Companyinfo = new CompanyInfo()
                                {
                                    CompanyName = parts[0].Trim('"'),
                                    AddressLine1 = parts[1].Trim('"'),
                                    AddressLine2 = parts[2].Trim('"'),
                                    County = parts[3].Trim('"'),
                                    PostalCode = parts[4].Trim('"')
                                };

                                // Add the CompanyInfo object to the list
                                companyList.Add(Companyinfo);
                            }
                        }
                        Console.WriteLine("Reading Data in Object");

                        // Read and display the stored CompanyInfo objects
                        foreach (CompanyInfo companyInfo in companyList)
                        {
                            Console.WriteLine("Company Name: " + companyInfo.CompanyName);
                            Console.WriteLine("Address Line 1: " + companyInfo.AddressLine1);
                            Console.WriteLine("Address Line 2: " + companyInfo.AddressLine2);
                            Console.WriteLine("County: " + companyInfo.County);
                            Console.WriteLine("Postal Code: " + companyInfo.PostalCode);
                            Console.WriteLine("---------------------------------------------");
                        }

                        // Create Word document
                        Console.WriteLine("Creating Doc. This can take up to 1 min...");
                        creatingwordfile.CreateWordDocument(companyList, "D:\\CompanyInfo.docx");
                        Console.WriteLine("Word document created successfully.");
                    }
                    catch (IOException e)
                    {
                        Debug.WriteLine("Error reading the file: " + e.Message);
                    }
                }
                else
                {
                    Console.WriteLine("Invalid file path or not a CSV file. Please try again.");
                }
            } while (!isFileValid);
            Debug.WriteLine("Created by Avi.");
            Debug.WriteLine("GitHub: ");
            Debug.WriteLine("Click enter to close.");
            Console.ReadLine();
        }
    }
}
