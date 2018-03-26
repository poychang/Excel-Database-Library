using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDatabaseLibrary
{
    public class Utilities
    {
        /// <summary>檢查檔案是否存在</summary>
        /// <param name="path">檔案路徑</param>
        /// <returns></returns>
        public static bool IsFileExist(string path) => File.Exists(path);

        /// <summary>檢查檔案是否為使用狀態</summary>
        /// <returns>是否為開啟狀態</returns>
        public static bool IsFileLocked(string storedPath)
        {
            try
            {
                using (File.Open(storedPath, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException exception)
            {
                var errorCode = Marshal.GetHRForException(exception) & 65535;
                return errorCode == 32 || errorCode == 33;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
    }
}
