using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PPPayReportTools.ExcelInterface
{
        public class ExcelCellExpressDeepUpdate<T> : IExcelCellExpressDeepUpdate<T>
        {
            private Regex cellPointRegex = new Regex("[A-Z]+[0-9]+");

            private Action<ICellModel> updateCellPointFunc { get; set; }
            public Func<T, bool> CheckContinuteFunc { get; set; }

            public ExcelCellExpressDeepUpdate(Action<ICellModel> updateCellPointFunc, Func<T, bool> checkIsContinuteFunc)
            {
                this.updateCellPointFunc = updateCellPointFunc;
                this.CheckContinuteFunc = checkIsContinuteFunc;
            }

            public bool IsContinute(T t)
            {
                return this.CheckContinuteFunc(t);
            }

            public string GetNextCellExpress(string currentExpress)
            {
                string nextCellExpress = currentExpress;

                List<ICellModel> cellModelList = this.GetCellModelList(currentExpress);
                string oldPointStr = null;
                string newPointStr = null;
                foreach (var item in cellModelList)
                {
                    oldPointStr = item.GetCellPosition();
                    this.updateCellPointFunc(item);
                    newPointStr = item.GetCellPosition();

                    nextCellExpress = nextCellExpress.Replace(oldPointStr, newPointStr);
                }
                return nextCellExpress;
            }


            private List<ICellModel> GetCellModelList(string cellExpress)
            {
                List<ICellModel> cellModelList = new List<ICellModel>(100);
                MatchCollection matchCollection = this.cellPointRegex.Matches(cellExpress);

                foreach (Match matchItem in matchCollection)
                {
                    cellModelList.Add(CellFactory.GetCellByExcelPosition(matchItem.Value));
                }
                return cellModelList;
            }

        }
}
