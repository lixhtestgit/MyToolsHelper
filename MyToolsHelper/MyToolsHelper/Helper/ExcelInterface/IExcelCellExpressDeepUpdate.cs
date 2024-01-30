using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.ExcelInterface
{
    public interface IExcelCellExpressDeepUpdate<T> : IExcelCellDeepUpdate
    {
        string GetNextCellExpress(string currentExpress);
        bool IsContinute(T t);

    }
}
