using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.ExcelInterface
{
    /// <summary>
    /// 单元格坐标深度更新接口
    /// </summary>
    public interface IExcelCellPointDeepUpdate : IExcelCellDeepUpdate
    {
        ICellModel GetNextCellPoint(ICellModel cellModel);
    }
}
