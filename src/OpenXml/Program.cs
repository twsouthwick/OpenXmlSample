// Licensed under the MIT license. See LICENSE file in the samples root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace OpenXml
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var filePath = Path.GetTempPath();

            CreateSpreadsheet(Path.Combine(filePath, "CustomersUsingStylesFromCode.xlsx"), OpenXmlExtensions.AddSpreadsheetStylesThroughCode);
            CreateSpreadsheet(Path.Combine(filePath, "CustomersUsingPreDefinedStyles.xlsx"), OpenXmlExtensions.AddPreDefinedSpreadsheetStyles);

            if (Debugger.IsAttached)
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Creates a spreadsheet on disk and adds styling and data to it.
        /// </summary>
        /// <param name="fileName">Path and file name</param>
        /// <param name="addStyleDelegate">This is delegate for adding the styles</param>
        private static void CreateSpreadsheet(string fileName, Action<WorkbookPart> addStyleDelegate)
        {
            try
            {
                using (var spreadsheet = SpreadsheetDocument.Create(fileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    spreadsheet.AddWorkbookPart();
                    spreadsheet.WorkbookPart.Workbook = new Workbook();

                    addStyleDelegate(spreadsheet.WorkbookPart);

                    AddCustomersSheet(spreadsheet);

                    spreadsheet.Save();

                    Console.WriteLine($"Created spreadsheet {fileName}");
                }
            }
            catch (IOException)
            {
                Console.WriteLine($"{fileName} could not be created. Check to see if it is already in use.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occured: {e.Message}");
            }
        }

        /// <summary>
        /// Adds a list of customers and header row to a spreadsheet
        /// </summary>
        private static void AddCustomersSheet(SpreadsheetDocument spreadsheetDocument)
        {
            var worksheet = spreadsheetDocument.AddWorksheet("Customers");

            // Add HeaderRow
            worksheet.AddRow(data: typeof(Customer).GetProperties(), styleIndex: 2);

            // Add DataRows
            foreach (var customer in GenerateCustomers())
            {
                var array = new object[]
                {
                    customer.Name,
                    customer.Address,
                    customer.City,
                    customer.State,
                    customer.ZipCode,
                    customer.DateEntered,
                };

                worksheet.AddRow(array, 1);
            }
        }

        /// <summary>
        /// Generates a collection of customers
        /// </summary>
        /// <param name="count">Number of items to add to collection</param>
        /// <returns>Collection of customers</returns>
        private static IEnumerable<Customer> GenerateCustomers(int count = 10)
        {
            for (var i = 0; i < count; i++)
            {
                yield return new Customer
                {
                    Name = $"Name {i}",
                    Address = $"Address {i}",
                    City = $"City {i}",
                    State = $"State {i}",
                    ZipCode = $"ZipCode {i}",
                    DateEntered = DateTimeOffset.Now
                };
            }
        }
    }
}
