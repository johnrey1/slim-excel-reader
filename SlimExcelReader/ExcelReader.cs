using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SlimExcelReader
{

    public class ExcelReader : IDisposable
    {
        private bool disposedValue;

        private Stream ExcelStream { get; set; }

        private SharedStringTable StringTable { get { return ExcelDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable; } }

        private SpreadsheetDocument ExcelDoc { get; set; }

        public string ExcelFilePath { get; set; }

        public string SheetName { get; set; }

        public ExcelReader(string excelFilePath)
        {
            this.ExcelFilePath = excelFilePath;
        }

        /// <summary>
        /// Opens the specified filepath for reading
        /// </summary>
        /// <returns></returns>
        public void OpenExcelReader()
        {
            if (ExcelDoc != null)
            {
                ExcelDoc.Close();
                ExcelDoc.Dispose();
            }

            if (ExcelStream != null)
            {
                ExcelStream.Close();
                ExcelStream.Dispose();
            }

            ExcelStream = System.IO.File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read);

            ExcelDoc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ExcelStream, false);
        }

        /// <summary>
        /// Gets the sheet data named in the SheetName property or the first sheet if no name specified
        /// </summary>
        /// <returns></returns>
        private SheetData GetSheetData()
        {
            if (ExcelDoc == null)
                OpenExcelReader();

            var workbookPart = ExcelDoc.WorkbookPart;
            var workbook = workbookPart.Workbook;

            var sheets = workbook.Descendants<Sheet>();

            if (string.IsNullOrEmpty(SheetName))
            {
                if (workbookPart.GetPartsOfType<WorksheetPart>().Count() > 0
                    && workbookPart.GetPartsOfType<WorksheetPart>().First().Worksheet.Elements<SheetData>().Count() > 0)
                    return workbookPart.GetPartsOfType<WorksheetPart>().First().Worksheet.Elements<SheetData>().First();
                else
                    throw new ApplicationException("No sheet data");
            }

            foreach (var sheet in sheets)
            {
                if (sheet.Name.Value.Equals(SheetName, StringComparison.CurrentCultureIgnoreCase))
                {
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    return worksheetPart.Worksheet.Elements<SheetData>().First();
                }
            }

            throw new ApplicationException(string.Format("{0} not found", SheetName));
        }

        /// <summary>
        /// Converts an Excel name for a cell into a 1-indexed row / column tuple
        /// </summary>
        /// <param name="cellId"></param>
        /// <returns></returns>
        public static Tuple<int, int> GetRowColumnFromCellId(string cellId)
        {
            var cellMatch = Regex.Match(cellId,@"([a-zA-Z]+)([\d]+)");

            var columnLettersReverse = cellMatch.Groups[1].ToString().ToUpper().Reverse();

            var column = 0;
            var counter = 0;
            foreach( var c in columnLettersReverse)
            {
                if (counter == 0)
                    column += (c - 64);
                else
                    column += (c - 64) * (26 * counter);
                counter++;
            }

            var row = int.Parse(cellMatch.Groups[2].ToString());
            
            return new Tuple<int, int>(row, column);
        }

        private static string GetCellValue(Cell cell, SharedStringTable ssTable)
        {
            if (int.TryParse(cell.InnerText, out int stringTableIndex))
                return ssTable.ElementAt(stringTableIndex).InnerText;
            else
                return null;
        }

        /// <summary>
        /// Returns 1-based index cell from spreadsheet
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>null if indices are out of bounds</returns>
        public string GetValue(int row, int column)
        {
            var sheetData = GetSheetData();

            if (row < 1 || column < 1)
                return null;
            try
            {
                var rowData = sheetData.Elements<Row>().Skip(row - 1).First();

                var cell = rowData.Elements<Cell>().Skip(column - 1).First();

                return GetCellValue(cell, StringTable);
            }
            catch(Exception)
            {
                return null;
            }
        }

        public string GetValue(string cellId)
        {
            var cellIndex = GetRowColumnFromCellId(cellId);

            return GetValue(cellIndex.Item1, cellIndex.Item2);
        } 

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    ExcelStream.Dispose();
                    ExcelDoc.Dispose();
                }

                ExcelStream = null;
                ExcelDoc = null;
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}