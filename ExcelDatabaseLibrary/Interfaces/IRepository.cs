using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelDatabaseLibrary.Interfaces
{
    /// <summary>資料操作介面</summary>
    /// <typeparam name="TEntity">資料表模型類別</typeparam>
    public interface IRepository<TEntity>
    {
        /// <summary>新增一筆資料</summary>
        /// <param name="entity">要新增的資料</param>
        void Create(TEntity entity);

        /// <summary>取得第一筆符合條件的資料</summary>
        /// <param name="predicate">取得資料的條件運算式</param>
        /// <returns>第一筆符合條件的資料</returns>
        TEntity Read(Expression<Func<TEntity, bool>> predicate);

        /// <summary>取得全部資料</summary>
        /// <returns>可查詢的資料</returns>
        IEnumerable<TEntity> Read();

        /// <summary>更新一筆資料</summary>
        /// <param name="entity">要被更新的資料</param>
        void Update(TEntity entity);

        /// <summary>刪除一筆資料</summary>
        /// <param name="entity">要被刪除的資料</param>
        void Delete(TEntity entity);
    }
}
