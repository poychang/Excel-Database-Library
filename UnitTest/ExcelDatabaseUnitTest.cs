using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelDatabaseLibrary;
using ExcelDatabaseLibrary.DataAnnotations;
using ExcelDatabaseLibrary.Extensions;
using UnitTest.Models;

namespace UnitTest
{
    [TestClass]
    public class ExcelDatabaseUnitTest
    {
        private static string GenerateStorePath() =>
            $"{AppDomain.CurrentDomain.BaseDirectory}ExcelDb-{Guid.NewGuid()}.xlsx";


        [TestMethod]
        public void ExcelDatabase_建立此實體時_建立ExcelDb資料檔()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式
            var storePath = GenerateStorePath();

            // Act: 呼叫目標物件的方法
            var excelDb = new ExcelDb(storePath);
            var isCreated = File.Exists(storePath);

            // Assert: 驗證是否符合預期
            Assert.IsTrue(isCreated);

            // Clearance UnitTest
            File.Delete(storePath);
        }

        [TestMethod]
        public void DeleteDatabase_呼叫時_刪除檔案資料庫()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式
            var storePath = GenerateStorePath();
            var excelDb = new ExcelDb(storePath);

            // Act: 呼叫目標物件的方法
            excelDb.DeleteDatabase();
            var isExist = File.Exists(storePath);

            // Assert: 驗證是否符合預期
            Assert.IsFalse(isExist);

            // Clearance UnitTest
            File.Delete(storePath);
        }

        [TestMethod]
        public void CreateTable_呼叫時_建立頁籤()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式
            var storePath = GenerateStorePath();
            var excelDb = new ExcelDb(storePath);

            // Act: 呼叫目標物件的方法
            excelDb.CreateTable<TableModel>();

            // Assert: 驗證是否符合預期
            using (var document = SpreadsheetDocument.Open(storePath, true))
            {
                var isExist = document.WorksheetIsExist(typeof(TableModel).GetAttribute<TableAttribute>().Name);
                Assert.IsTrue(isExist);
            }

            // Clearance UnitTest
            File.Delete(storePath);
        }

        [TestMethod]
        public void DeleteTable_呼叫時_刪除頁籤()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式
            var storePath = GenerateStorePath();
            var excelDb = new ExcelDb(storePath);
            excelDb.CreateTable<TableModel>();

            // Act: 呼叫目標物件的方法
            excelDb.DeleteTable<TableModel>();

            // Assert: 驗證是否符合預期
            using (var document = SpreadsheetDocument.Open(storePath, true))
            {
                var isExist = document.WorksheetIsExist(typeof(TableModel).GetAttribute<TableAttribute>().Name);
                Assert.IsFalse(isExist);
            }

            // Clearance UnitTest
            File.Delete(storePath);
        }
    }
}