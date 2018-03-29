using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDatabaseLibrary.DataAnnotations;
using ExcelDatabaseLibrary.Extensions;
using ExcelDatabaseLibrary.Interfaces;
using Microsoft.Extensions.Options;

namespace ExcelDatabaseLibrary
{
    public class ExcelDb : IDocumentDb
    {
        /// <inheritdoc />
        public string StoredPath { get; }

        private bool _isCreated;

        /// <inheritdoc />
        public bool IsCreated
        {
            get
            {
                _isCreated = Utilities.IsFileExist(StoredPath);
                return _isCreated;
            }
        }

        /// <summary>建構式</summary>
        /// <param name="storedPath">檔案資料庫的儲存路徑</param>
        public ExcelDb(string storedPath)
        {
            StoredPath = storedPath;
            _isCreated = CreateDatabase();
        }

        /// <summary>建構式</summary>
        /// <param name="optionsAccessor">選項存取子</param>
        public ExcelDb(IOptions<ExcelDbOptions> optionsAccessor)
        {
            StoredPath = optionsAccessor.Value.StoredPath;
            _isCreated = CreateDatabase();
        }

        /// <inheritdoc />
        public bool CreateDatabase()
        {
            const string tablename = "master";
            try
            {
                if (Utilities.IsFileExist(StoredPath)) return true;

                var memoryStream = new MemoryStream();
                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ??
                                 workbookPart.Workbook.AppendChild(new Sheets());

                    if (document.WorksheetIsExist(tablename)) return false;

                    var sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = document.NextWorksheetSerialNumber(),
                        Name = tablename
                    };
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();
                    document.Close();

                    using (var fileStream = new FileStream(StoredPath, FileMode.Create))
                    {
                        memoryStream.WriteTo(fileStream);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <inheritdoc />
        public bool DeleteDatabase()
        {
            try
            {
                if (!Utilities.IsFileExist(StoredPath)) return true;

                File.Delete(StoredPath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <inheritdoc />
        public bool CreateTable<T>(string tableName = null)
        {
            tableName = FetchTableName<T>(tableName);

            try
            {
                using (var document = SpreadsheetDocument.Open(StoredPath, true))
                {
                    if (document.WorksheetIsExist(tableName)) return false;

                    var workbookPart = document.WorkbookPart;

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var rowOfColumnName = new Row(FetchColumnField(typeof(T)));
                    worksheetPart.Worksheet = new Worksheet(new SheetData(rowOfColumnName));

                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ??
                                 workbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = document.NextWorksheetSerialNumber(),
                        Name = tableName
                    };
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();
                    document.Close();

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <inheritdoc />
        public bool DeleteTable<T>(string tableName = null)
        {
            tableName = FetchTableName<T>(tableName);

            try
            {
                using (var document = SpreadsheetDocument.Open(StoredPath, true))
                {
                    var workbookPart = document.WorkbookPart;
                    var worksheetPart = document.GetWorksheetPart(tableName);

                    // Remove the sheet reference from the workbook.
                    document.GetSheet(tableName).Remove();
                    // Delete the worksheet part.
                    workbookPart.DeletePart(worksheetPart);

                    workbookPart.Workbook.Save();
                    document.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <summary>取得資料表名稱</summary>
        /// <typeparam name="T">資料表模型類別</typeparam>
        /// <param name="tableName">資料表名稱</param>
        /// <returns></returns>
        private static string FetchTableName<T>(string tableName)
        {
            return string.IsNullOrEmpty(tableName) ? typeof(T).GetAttribute<TableAttribute>().Name : tableName;
        }

        /// <summary>取得資料表欄位清單</summary>
        /// <param name="tableType">資料表模型類別</param>
        /// <returns></returns>
        private static IEnumerable<Cell> FetchColumnField(Type tableType)
        {
            var columnNameList = new List<Cell>();
            foreach (var property in tableType.GetProperties())
            {
                var columnName = tableType.GetAttributeProperty<ColumnAttribute>(property.Name)?.Name;
                var item = string.IsNullOrEmpty(columnName) ? property.Name : columnName;
                columnNameList.Add(new Cell() {CellValue = new CellValue(item), DataType = CellValues.String});
            }

            return columnNameList;
        }
    }
}