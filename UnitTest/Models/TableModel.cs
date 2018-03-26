using System;
using ExcelDatabaseLibrary.DataAnnotations;

namespace UnitTest.Models
{
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
