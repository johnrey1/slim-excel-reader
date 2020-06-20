using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace SlimExcelReader.Tests
{
    [TestClass()]
    public class SlimExcelReaderTests
    {
        private static readonly Dictionary<string, string> lookups = new Dictionary<string, string>(new KeyValuePair<string, string>[]
            { new KeyValuePair<string, string>("O41","9Ok$4Sx,"),
                new KeyValuePair<string, string>("BP183","8Xp&6Sd/")
            });

        private readonly string ExcelFile = "FatExcel.xlsx";
        private ExcelReader ExcelReader = null;
        private readonly string ValueSheetName = "Value 1";
        private readonly string FormulaSheetName = "Value 2";
        private readonly KeyValuePair<string, string> kvpFormulaSheet = new KeyValuePair<string, string>("A1", "5Rg+7Vd*");

        [TestInitialize]
        public void Initialize()
        {
            if(!File.Exists(ExcelFile))
                ZipFile.ExtractToDirectory(ExcelFile + ".zip", System.Environment.CurrentDirectory, true);
            if (ExcelReader == null)
            {
                ExcelReader = new ExcelReader(ExcelFile);
                ExcelReader.OpenExcelReader();
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            ExcelReader.Dispose();
        }

        [TestMethod]
        public void GetCellExistTest()
        {            
            ExcelReader.SheetName = ValueSheetName;            

            var result = ExcelReader.GetValue("O41");
            Assert.AreEqual(lookups["O41"], result);

            result = ExcelReader.GetValue("BP183");
            Assert.AreEqual(lookups["BP183"], result);
        }

        [TestMethod]
        public void GetSheetDoesNotExist()
        {
            ExcelReader.SheetName = "InvalidSheetName";

            Assert.ThrowsException<ApplicationException>(() => ExcelReader.GetValue("O41"), ExcelReader.SheetName + " not found");
        }

        [TestMethod]
        public void GetCellDoesNotExistTest()
        {
            ExcelReader.SheetName = ValueSheetName;
            
            var result = ExcelReader.GetValue("YYY1123");

            Assert.IsNull(result);
        }

        /// <summary>
        /// Currently do not support formula cells and return null
        /// </summary>
        [TestMethod]
        public void GetValueFromFormulaCell()
        {
            ExcelReader.SheetName = FormulaSheetName;
            
            var result = ExcelReader.GetValue(kvpFormulaSheet.Key);
            
            Assert.IsNull(result);
        }
        
        [TestMethod]
        public void FileDoesNotExistTest()
        {
            var excelReader = new ExcelReader("ThisFileDoesntExist");
            Assert.ThrowsException<FileNotFoundException>(() => excelReader.OpenExcelReader());
        }

        [TestMethod()]
        public void GetRowColumnFromCellIdTest()
        {
            var cell1 = "XEA1293";
            var cell1ColumnExpected = 1379;
            var cell1RowExpected = 1293;

            var result = ExcelReader.GetRowColumnFromCellId(cell1);

            Assert.AreEqual(cell1RowExpected, result.Item1);
            Assert.AreEqual(cell1ColumnExpected, result.Item2);
        }
    }
}