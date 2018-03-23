using System;
using System.Runtime.CompilerServices;

namespace ExcelDatabaseLibrary.DataAnnotations
{
    /// <summary>資料表的屬性驗證模型</summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public sealed class TableAttribute : Attribute
    {
        /// <summary>資料表名稱</summary>
        public string Name { get; set; }

        /// <summary>建構式</summary>
        /// <param name="callerMemberName">調用者的名稱</param>
        public TableAttribute([CallerMemberName] string callerMemberName = null)
        {
            Name = string.IsNullOrEmpty(callerMemberName) ? new Guid().ToString() : callerMemberName;
        }
    }
}