using PPPayReportTools.ExcelInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.Excel
{
    public class ExcelFileDescription
    {
        public ExcelFileDescription(int titleRowIndex)
        {
            this.TitleRowIndex = titleRowIndex;
        }

        public ExcelFileDescription(IExcelDeepUpdate excelDeepUpdate)
        {
            this.ExcelDeepUpdateList = new List<IExcelDeepUpdate> { excelDeepUpdate };
        }
        public ExcelFileDescription(List<IExcelDeepUpdate> excelDeepUpdateList)
        {
            this.ExcelDeepUpdateList = excelDeepUpdateList;
        }

        /// <summary>
        /// 标题所在行位置（默认为0，没有标题填-1）
        /// </summary>
        public int TitleRowIndex { get; set; }

        /// <summary>
        /// Excel深度更新策略
        /// </summary>
        public List<IExcelDeepUpdate> ExcelDeepUpdateList { get; set; }

    }
}
