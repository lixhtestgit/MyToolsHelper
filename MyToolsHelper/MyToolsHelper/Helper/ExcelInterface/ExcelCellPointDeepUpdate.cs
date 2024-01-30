using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPPayReportTools.ExcelInterface
{
    public class ExcelCellPointDeepUpdate : IExcelCellPointDeepUpdate
    {
        private Action<ICellModel> updateCellPointFunc { get; set; }


        public ExcelCellPointDeepUpdate(Action<ICellModel> updateCellPointFunc)
        {
            this.updateCellPointFunc = updateCellPointFunc;
        }

        public ICellModel GetNextCellPoint(ICellModel cellModel)
        {
            ICellModel nextCell = null;

            ICellModel cell = new CellModel(cellModel.RowIndex, cellModel.ColumnIndex);
            if (cellModel != null && this.updateCellPointFunc != null)
            {
                this.updateCellPointFunc(cell);
                if (cell.RowIndex != cellModel.RowIndex || cell.ColumnIndex != cellModel.ColumnIndex)
                {
                    nextCell = cell;
                }
            }

            return nextCell;
        }

    }
}
