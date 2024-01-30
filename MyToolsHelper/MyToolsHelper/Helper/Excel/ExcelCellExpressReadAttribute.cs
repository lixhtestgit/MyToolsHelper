using PPPayReportTools.ExcelInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.Excel
{
    /// <summary>
    /// Excel单元格-表达式读取-标记特性
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Field | System.AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelCellExpressReadAttribute : System.Attribute
    {
        /// <summary>
        /// 读取数据使用：该参数使用表达式生成数据（Excel文件中支持的表达式均可以，可以是单元格位置也可以是表达式（如：A1,B2,C1+C2...））
        /// </summary>
        public string CellCoordinateExpress { get; set; }

        /// <summary>
        /// 字符输出格式（数字和日期类型需要）
        /// </summary>
        public string OutputFormat { get; set; }

        /// <summary>
        /// 生成单元格表达式读取特性
        /// </summary>
        /// <param name="cellCoordinateExpress">初始单元格表达式</param>
        /// <param name="outputFormat">（可选）格式化字符串</param>
        public ExcelCellExpressReadAttribute(string cellCoordinateExpress, string outputFormat = "")
        {
            this.CellCoordinateExpress = cellCoordinateExpress;
            this.OutputFormat = outputFormat;
        }
    }
}
