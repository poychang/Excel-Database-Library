using System;
using System.Linq;
using ExcelDatabaseLibrary.DataAnnotations;
using ExcelDatabaseLibrary.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest
{
    [TestClass]
    public class TypeExtensionUnitTest
    {
        [TestMethod]
        public void GetAttribute_呼叫時指定要取得的屬性驗證模型_取回屬性驗證模型實體()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式

            // Act: 呼叫目標物件的方法
            var tableAttribute = typeof(TestModel).GetAttribute<TableAttribute>();

            // Assert: 驗證是否符合預期
            Assert.AreEqual(tableAttribute.Name, "我的資料表");
        }

        [TestMethod]
        public void GetAttributeProperty_呼叫時指定要取得的屬性驗證模型及其屬性名稱_取回屬性驗證模型實體()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式

            // Act: 呼叫目標物件的方法
            var result = typeof(TestModel).GetAttributeProperty<ColumnAttribute>(nameof(TestModel.Id));

            // Assert: 驗證是否符合預期
            Assert.AreEqual(result.Name, "識別碼");
        }

        [TestMethod]
        public void GetAttributeProperties_呼叫時指定要取得的屬性驗證模型_取回其所有屬性的屬性驗證模型清單()
        {
            // Arrange: 初始化目標物件、相依物件、方法參數、預期結果，或是預期與相依物件的互動方式

            // Act: 呼叫目標物件的方法
            var properties = typeof(TestModel).GetAttributeProperties<ColumnAttribute>();

            // Assert: 驗證是否符合預期
            Assert.AreEqual(properties.Count(), 2);
        }
    }

    [Table(Name = "我的資料表")]
    internal class TestModel
    {
        [Column(Name = "識別碼")]
        public string Id { get; set; }
        [Column(Name = "建立日期")]
        public DateTimeOffset CreateDate { get; set; }
    }
}
