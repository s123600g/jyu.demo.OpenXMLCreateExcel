using System.Collections.Generic;

namespace ExcelGenerator.Models
{
    public class ExcelContentEntity
    {
        /// <summary>
        /// 表單清單
        /// </summary>
        public List<ExcelSheetEntity> Sheets { get; set; } = new List<ExcelSheetEntity>();
    }

    public class ExcelSheetEntity
    {
        /// <summary>
        /// 表單名稱
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 表單資料
        /// </summary>
        public List<List<dynamic>> RowData { get; set; } = new List<List<dynamic>>();
    }
}
