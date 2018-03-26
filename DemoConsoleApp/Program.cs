using System;
using ExcelDatabaseLibrary;
using ExcelDatabaseLibrary.DataAnnotations;
using ExcelDatabaseLibrary.Extensions;

namespace DemoConsoleApp
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // 建立 Excel 資料庫
            var excelDb = new ExcelDb($"{AppDomain.CurrentDomain.BaseDirectory}ExcelDb.xlsx");
            //Console.WriteLine($"Excel Database is existed: {excelDb.IsCreated}");

            // 刪除 Excel 資料庫
            //excelDb.DeleteDatabase();
            //Console.WriteLine($"Excel Database is existed: {excelDb.IsCreated}");

            // 列出資料表名稱及其欄位名稱
            //Console.WriteLine($"Table Name: {typeof(TableModel).GetAttribute<TableAttribute>().Name}");
            //Console.WriteLine("Column List: ");
            //foreach (var property in typeof(TableModel).GetAttributeProperties<ColumnAttribute>())
            //{
            //    Console.WriteLine(property.Name);
            //}

            Console.WriteLine(excelDb.CreateTable<TableModel>());
            //Console.WriteLine(excelDb.DeleteTable<TableModel>());
            Console.ReadLine();
        }
    }

    [Table(Name = "我的資料表")]
    internal class TableModel
    {
        [Column(Name = "識別碼")]
        public string Id { get; set; }

        [Column(Name = "姓名")]
        public string Name { get; set; }

        [Column(Name = "年齡")]
        public int Age { get; set; }

        [Column(Name = "建立日期")]
        public DateTimeOffset CreateDate { get; set; }
    }
}