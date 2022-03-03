using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;

namespace ExcelLoader
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = $"{GetCurrentFolder()}\\Test.xlsx";
            var sheetName = "sheet";

            Console.WriteLine("Writing Excel file");
            WriteExcelFileUsingOpenXML(file, sheetName);

            Console.WriteLine("Reading Excel data");
            var table = ReadExcellDataUsingOpenXML(file, sheetName);

            foreach (var row in table.Rows)
            {
                var dataRow = (DataRow)row;
                Console.WriteLine($"{dataRow[0]} - {dataRow[1]} - {dataRow[2]}");
            }

            Console.ReadLine();
        }

        private static DataTable ReadExcellDataUsingOdbc(string fileName, string sheetName)
        {
            string strConnString = $"Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};Dbq={fileName};Extensions=xls/xlsx;Persist Security Info=False";
            DataTable dataTable;
            using (OdbcConnection oConn = new OdbcConnection(strConnString))
            {
                using (OdbcCommand oCmd = new OdbcCommand())
                {
                    oCmd.Connection = oConn;

                    oCmd.CommandType = System.Data.CommandType.Text;
                    oCmd.CommandText = "select * from [" + sheetName + "$]";

                    OdbcDataAdapter oAdap = new OdbcDataAdapter();
                    oAdap.SelectCommand = oCmd;

                    dataTable = new DataTable();
                    oAdap.Fill(dataTable);
                    oAdap.Dispose();
                }
            }

            return dataTable;
        }

        private static DataTable ReadExcellDataUsingOpenXML(string fileName, string sheetName)
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Test 1");
            dataTable.Columns.Add("Test 2");
            dataTable.Columns.Add("Test 3");

            using (SpreadsheetDocument spreadsheetDocument =
                SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                foreach (var row in rows)
                {
                    var cells = row.Elements<Cell>();
                    var columnIndex = 0;
                    var dicValues = new Dictionary<int, string>();
                    foreach (var cell in cells)
                    {
                        dicValues[columnIndex] = cell.CellValue.Text;
                        columnIndex++;
                    }

                    var newRow = dataTable.NewRow();
                    newRow[0] = dicValues[0];
                    newRow[1] = dicValues[1];
                    newRow[2] = dicValues[2];

                    dataTable.Rows.Add(newRow);
                }
            }

            return dataTable;
        }

        //Obs: Using OpenXML it is not possible to load file content using ODBC
        private static void WriteExcelFileUsingOpenXML(string FileName, string sheetName)
        {
            var persons = GetPersons();

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }

        public static void WriteExcelFileUsingCloseXML(string FileName, string sheetName)
        {
            var persons = GetPersons();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(sheetName);
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Name";
                worksheet.Cell(currentRow, 2).Value = "Last Name";
                worksheet.Cell(currentRow, 3).Value = "Age";
                foreach (var person in persons)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = person.FirstName;
                    worksheet.Cell(currentRow, 2).Value = person.LastName;
                    worksheet.Cell(currentRow, 3).Value = person.Age;
                }

                workbook.SaveAs(FileName);
            }
        }

        private static List<Person> GetPersons()
        {
            List<Person> persons = new List<Person>()
            {
                new Person() {FirstName="Vaan", LastName="No last name", Age = 15},
                new Person() {FirstName="Basch", LastName="From Ronsenburg", Age = 30}
            };

            for (int i = 0; i < 10; i++)
            {
                persons.Add(new Person() { FirstName = $"More {i}", LastName="Last name", Age = 20 + i });
            }

            return persons;
        }

        private static string GetCurrentFolder()
        {
            return Environment.CurrentDirectory;
        }
    }
}