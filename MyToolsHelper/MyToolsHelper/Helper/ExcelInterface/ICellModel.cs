using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.ExcelInterface
{
    public interface ICellModel
    {
        int RowIndex { get; set; }
        int ColumnIndex { get; set; }
        object CellValue { get; set; }

        bool IsCellFormula { get; set; }

        string GetCellPosition();

    }
}
