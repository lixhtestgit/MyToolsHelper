using NPOI.SS.UserModel;
using PPPayReportTools.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace PPPayReportTools.ExcelInterface
{
    /// <summary>
    /// 单元格工厂类
    /// </summary>
    public class CellFactory
    {
        private static Regex _CellPostionRegex = new Regex("[A-Z]+\\d+");
        private static Regex _RowRegex = new Regex("\\d+");

        /// <summary>
        /// 通过Excel单元格坐标位置初始化对象
        /// </summary>
        /// <param name="excelCellPosition">A1,B2等等</param>
        /// <returns></returns>
        public static ICellModel GetCellByExcelPosition(string excelCellPosition)
        {
            CellModel cellModel = null;

            bool isMatch = CellFactory._CellPostionRegex.IsMatch(excelCellPosition);
            if (isMatch)
            {
                Match rowMath = CellFactory._RowRegex.Match(excelCellPosition);
                int rowPositon = Convert.ToInt32(rowMath.Value);
                int rowIndex = rowPositon - 1;
                int columnIndex = CellFactory.GetExcelColumnIndex(excelCellPosition.Replace(rowPositon.ToString(), ""));

                cellModel = new CellModel(rowIndex, columnIndex);
            }
            return cellModel;
        }

        /// <summary>
        /// 将数据放入单元格中
        /// </summary>
        /// <param name="cell">单元格对象</param>
        /// <param name="cellValue">数据</param>
        /// <param name="outputFormat">格式化字符串</param>
        /// <param name="isCoordinateExpress">是否是表达式数据</param>
        public static void SetCellValue(ICell cell, object cellValue, string outputFormat, bool isCoordinateExpress)
        {
            if (cell != null)
            {
                if (isCoordinateExpress)
                {
                    cell.SetCellFormula(cellValue.ToString());
                }
                else
                {
                    if (!string.IsNullOrEmpty(outputFormat))
                    {
                        string formatValue = null;
                        IFormatProvider formatProvider = null;
                        if (cellValue is DateTime)
                        {
                            formatProvider = new DateTimeFormatInfo();
                            ((DateTimeFormatInfo)formatProvider).ShortDatePattern = outputFormat;
                        }
                        formatValue = ((IFormattable)cellValue).ToString(outputFormat, formatProvider);

                        cell.SetCellValue(formatValue);
                    }
                    else
                    {
                        if (cellValue is decimal || cellValue is double || cellValue is int)
                        {
                            cell.SetCellValue(Convert.ToDouble(cellValue));
                        }
                        else if (cellValue is DateTime)
                        {
                            cell.SetCellValue((DateTime)cellValue);
                        }
                        else if (cellValue is bool)
                        {
                            cell.SetCellValue((bool)cellValue);
                        }
                        else
                        {
                            cell.SetCellValue(cellValue.ToString());
                        }
                    }
                }
            }

        }

        public static void SetDeepUpdateCellValue(ISheet sheet, int rowIndex, int columnIndex, object cellValue, string outputFormat, bool isCoordinateExpress, List<IExcelCellPointDeepUpdate> excelDeepUpdateList)
        {
            if (sheet != null)
            {
                //更新起始单元格数据
                ICell nextCell = ExcelHelper.GetOrCreateCell(sheet, rowIndex, columnIndex);
                CellFactory.SetCellValue(nextCell, cellValue, outputFormat, isCoordinateExpress);

                #region 执行单元格深度更新策略

                ICellModel startCellPosition = new CellModel
                {
                    RowIndex = rowIndex,
                    ColumnIndex = columnIndex
                };

                ICellModel nextCellPosition = null;
                Action<IExcelCellPointDeepUpdate> actionDeepUpdateAction = (excelDeepUpdate) =>
                {
                    //获取起始执行单元格位置
                    nextCellPosition = excelDeepUpdate.GetNextCellPoint(startCellPosition);

                    //执行深度更新，一直到找不到下个单元格为止
                    do
                    {
                        nextCell = ExcelHelper.GetOrCreateCell(sheet, nextCellPosition.RowIndex, nextCellPosition.ColumnIndex);
                        if (nextCell != null)
                        {
                            CellFactory.SetCellValue(nextCell, cellValue, outputFormat, isCoordinateExpress);
                            nextCellPosition = excelDeepUpdate.GetNextCellPoint(nextCellPosition);
                        }
                    } while (nextCell != null);
                };

                foreach (var excelDeepUpdate in excelDeepUpdateList)
                {
                    actionDeepUpdateAction(excelDeepUpdate);
                }

                #endregion

            }

        }


        /// <summary>
        /// 数字转字母
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static string GetExcelColumnPosition(int number)
        {
            var a = number / 26;
            var b = number % 26;

            if (a > 0)
            {
                return CellFactory.GetExcelColumnPosition(a - 1) + (char)(b + 65);
            }
            else
            {
                return ((char)(b + 65)).ToString();
            }
        }

        /// <summary>
        /// 字母转数字
        /// </summary>
        /// <param name="columnPosition"></param>
        /// <returns></returns>
        public static int GetExcelColumnIndex(string zm)
        {
            int index = 0;
            char[] chars = zm.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index - 1;
        }

    }
}
