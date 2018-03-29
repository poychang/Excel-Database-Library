using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDatabaseLibrary.Extensions
{
    public static class OpenXmlExtension
    {
        /// <summary>取得指定頁籤名稱的 WorksheetPart</summary>
        /// <param name="document">SpreadsheetDocument 實體</param>
        /// <param name="sheetname">頁籤名稱</param>
        /// <returns></returns>
        public static WorksheetPart GetWorksheetPart(this SpreadsheetDocument document, string sheetname = null)
        {
            var sheet = document.WorkbookPart.Workbook.Descendants<Sheet>().SingleOrDefault(p => p.Name == sheetname);
            return (WorksheetPart) document.WorkbookPart.GetPartById(sheet.Id);
        }

        /// <summary>取得指定頁籤名稱的 Sheet</summary>
        /// <param name="document">SpreadsheetDocument 實體</param>
        /// <param name="sheetname">頁籤名稱</param>
        /// <returns></returns>
        public static Sheet GetSheet(this SpreadsheetDocument document, string sheetname = null)
        {
            return document.WorkbookPart.Workbook.Descendants<Sheet>().SingleOrDefault(p => p.Name == sheetname);
        }

        /// <summary>Worksheet 是否已存在</summary>
        /// <param name="document">SpreadsheetDocument 實體</param>
        /// <param name="sheetname">頁籤名稱</param>
        /// <returns></returns>
        public static bool WorksheetIsExist(this SpreadsheetDocument document, string sheetname)
        {
            return document.WorkbookPart.Workbook.Descendants<Sheet>().Any(p => p.Name == sheetname);
        }

        /// <summary>取得下一個頁籤識別序號</summary>
        /// <param name="document">SpreadsheetDocument 實體</param>
        /// <returns></returns>
        public static uint NextWorksheetSerialNumber(this SpreadsheetDocument document)
        {
            var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            return sheets.Descendants<Sheet>().Any()
                ? sheets.Elements<Sheet>().Select(p => p.SheetId.Value).Max() + 1
                : 1;
        }
    }
}