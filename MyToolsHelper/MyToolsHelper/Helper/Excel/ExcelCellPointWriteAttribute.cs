using PPPayReportTools.ExcelInterface;
using System;

namespace PPPayReportTools.Excel
{
    /// <summary>
    /// Excel单元格-表达式读取-标记特性
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Field | System.AttributeTargets.Property, AllowMultiple = true)]
    public class ExcelCellPointWriteAttribute : System.Attribute
    {
        /// <summary>
        /// 单元格位置（A3,B4...）
        /// </summary>
        public string CellPosition { get; set; }

        /// <summary>
        /// 字符输出格式（数字和日期类型需要）
        /// </summary>
        public string OutputFormat { get; set; }


        public ExcelCellPointWriteAttribute(string cellPosition, string outputFormat = null)
        {
            CellPosition = cellPosition;
            OutputFormat = outputFormat;
        }
    }
}
