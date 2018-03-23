using System;
using ExcelDatabaseLibrary.DataAnnotations;
using ExcelDatabaseLibrary.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest
{
    [TestClass]
    public class TypeExtensionUnitTest
    {
        [TestMethod]
        public void GetAttribute_指定類別及屬性驗證模型_取回類別的屬性驗證模型實體()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式

            // Act: 呼叫目標物件的方法
            var tableAttribute = typeof(TableModel).GetAttribute<TableAttribute>();

            // Assert: 驗證是否符合預期
            Assert.AreEqual(tableAttribute.Name, "我的資料表");
        }
    }

    [Table(Name = "我的資料表")]
    internal class TableModel
    {
        [Column(Name = "識別碼")]
        public string Id { get; set; }
        [Column(Name = "建立日期")]
        public DateTimeOffset CreateDate { get; set; }
    }
}
