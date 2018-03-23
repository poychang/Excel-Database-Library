using System;
using System.Runtime.CompilerServices;

namespace ExcelDatabaseLibrary.DataAnnotations
{
    /// <summary>資料表欄位的屬性驗證模型</summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public sealed class ColumnAttribute : Attribute
    {
        /// <summary>欄位名稱</summary>
        public string Name { get; set; }

        /// <summary>建構式</summary>
        /// <param name="callerMemberName">調用者的名稱</param>
        public ColumnAttribute([CallerMemberName] string callerMemberName = null)
        {
            Name = string.IsNullOrEmpty(callerMemberName) ? new Guid().ToString() : callerMemberName;
        }
    }
}