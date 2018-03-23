namespace ExcelDatabaseLibrary.Interfaces
{
    /// <summary>檔案資料庫介面</summary>
    public interface IDocumentDb
    {
        /// <summary>檔案資料庫的儲存路徑</summary>
        string StoredPath { get; set; }

        /// <summary>資料庫是否存在</summary>
        /// <returns>是否存在</returns>
        bool IsCreated();

        /// <summary>資料庫是否成功連線</summary>
        /// <returns>是否連線</returns>
        bool IsConnected();

        /// <summary>建立資料庫</summary>
        /// <returns>是否成功</returns>
        bool CreateDatabase();

        /// <summary>刪除資料庫</summary>
        /// <returns>是否成功</returns>
        bool DeleteDatabase();

        /// <summary>建立資料表</summary>
        /// <typeparam name="T">資料表模型類別</typeparam>
        /// <returns>是否成功</returns>
        bool CreateTable<T>(string tablename);

        /// <summary>刪除資料表</summary>
        /// <typeparam name="T">資料表模型類別</typeparam>
        /// <returns>是否成功</returns>
        bool DeleteTable<T>();
    }
}