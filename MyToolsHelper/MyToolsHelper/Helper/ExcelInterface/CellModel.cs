using System.Collections.Generic;
using System.Linq;

namespace PPPayReportTools.ExcelInterface
{
    public class CellModel : ICellModel
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public object CellValue { get; set; }

        public bool IsCellFormula { get; set; }

        public CellModel() { }

        /// <summary>
        /// 默认初始化对象
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="cellValue"></param>
        public CellModel(int rowIndex, int columnIndex, object cellValue = default(object)) : this(rowIndex, columnIndex, cellValue, false)
        {
        }

        /// <summary>
        /// 默认初始化对象
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="cellValue"></param>
        /// <param name="isCellFormula"></param>
        public CellModel(int rowIndex, int columnIndex, object cellValue, bool isCellFormula)
        {
            this.RowIndex = rowIndex;
            this.ColumnIndex = columnIndex;
            this.CellValue = cellValue;
            this.IsCellFormula = isCellFormula;
        }

        /// <summary>
        /// 获取单元格位置
        /// </summary>
        /// <returns></returns>
        public string GetCellPosition()
        {
            return CellFactory.GetExcelColumnPosition(this.ColumnIndex) + (this.RowIndex + 1).ToString();
        }
    }

    public class CellModelColl : List<CellModel>, IList<CellModel>
    {
        public CellModelColl() { }
        public CellModelColl(int capacity) : base(capacity)
        {

        }

        /// <summary>
        /// 根据行下标，列下标获取单元格数据
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public CellModel this[int rowIndex, int columnIndex]
        {
            get
            {
                CellModel cell = this.FirstOrDefault(m => m.RowIndex == rowIndex && m.ColumnIndex == columnIndex);
                return cell;
            }
            set
            {
                CellModel cell = this.FirstOrDefault(m => m.RowIndex == rowIndex && m.ColumnIndex == columnIndex);
                if (cell != null)
                {
                    cell.CellValue = value.CellValue;
                }
            }
        }

        public CellModel CreateOrGetCell(int rowIndex, int columnIndex)
        {
            CellModel cellModel = this[rowIndex, columnIndex];
            if (cellModel == null)
            {
                cellModel = new CellModel()
                {
                    RowIndex = rowIndex,
                    ColumnIndex = columnIndex
                };
                this.Add(cellModel);
            }
            return cellModel;
        }

        public CellModel GetCell(string cellStringValue)
        {
            CellModel cellModel = null;

            cellModel = this.FirstOrDefault(m => m.CellValue.ToString().Equals(cellStringValue, System.StringComparison.OrdinalIgnoreCase));

            return cellModel;
        }

        /// <summary>
        /// 所有一行所有的单元格数据
        /// </summary>
        /// <param name="rowIndex">行下标</param>
        /// <returns></returns>
        public List<CellModel> GetRawCellList(int rowIndex)
        {
            List<CellModel> cellList = null;
            cellList = this.FindAll(m => m.RowIndex == rowIndex);

            return cellList ?? new List<CellModel>(0);
        }

        /// <summary>
        /// 所有一列所有的单元格数据
        /// </summary>
        /// <param name="columnIndex">列下标</param>
        /// <returns></returns>
        public List<CellModel> GetColumnCellList(int columnIndex)
        {
            List<CellModel> cellList = null;
            cellList = this.FindAll(m => m.ColumnIndex == columnIndex);

            return cellList ?? new List<CellModel>(0);
        }

    }

}
