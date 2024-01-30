using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.Excel
{
    /// <summary>
    /// Excel单元格-模板参数写入-标记特性
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Field | System.AttributeTargets.Property, AllowMultiple = true)]
    public class ExcelCellParamWriteAttribute : System.Attribute
    {
        /// <summary>
        /// 模板文件的预定义变量使用（{A} {B}）
        /// </summary>
        public string CellParamName { get; set; }

        /// <summary>
        /// 字符输出格式（数字和日期类型需要）
        /// </summary>
        public string OutputFormat { get; set; }

        public ExcelCellParamWriteAttribute(string cellParamName, string outputFormat = "")
        {
            CellParamName = cellParamName;
            OutputFormat = outputFormat;
        }


    }
}
