using NLog;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PPPayReportTools.ExcelInterface;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PPPayReportTools.Excel
{
    /// <summary>
    /// EXCEL帮助类
    /// </summary>
    /// <typeparam name="T">泛型类</typeparam>
    /// <typeparam name="TCollection">泛型类集合</typeparam>
    public class ExcelHelper
    {
        private static Logger _Logger = LogManager.GetCurrentClassLogger();


        public static IWorkbook GetExcelWorkbook(string filePath)
        {
            IWorkbook workbook = null;

            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    try
                    {
                        workbook = new XSSFWorkbook(fileStream);
                    }
                    catch (Exception)
                    {
                        workbook = new HSSFWorkbook(fileStream);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception($"文件：{filePath}被占用！", e);
            }
            return workbook;
        }

        public static ISheet GetExcelWorkbookSheet(IWorkbook workbook, int sheetIndex = 0)
        {
            ISheet sheet = null;

            if (workbook != null)
            {
                if (sheetIndex >= 0)
                {
                    sheet = workbook.GetSheetAt(sheetIndex);
                }
            }
            return sheet;
        }

        public static ISheet GetExcelWorkbookSheet(IWorkbook workbook, string sheetName = "sheet1")
        {
            ISheet sheet = null;

            if (workbook != null && !string.IsNullOrEmpty(sheetName))
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
            }
            return sheet;
        }

        public static IRow GetOrCreateRow(ISheet sheet, int rowIndex)
        {
            IRow row = null;
            if (sheet != null)
            {
                row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    row = sheet.CreateRow(rowIndex);
                }
            }
            return row;
        }

        public static ICell GetOrCreateCell(ISheet sheet, int rowIndex, int columnIndex)
        {
            ICell cell = null;

            IRow row = ExcelHelper.GetOrCreateRow(sheet, rowIndex);
            if (row != null)
            {
                cell = row.GetCell(columnIndex);
                if (cell == null)
                {
                    cell = row.CreateCell(columnIndex);
                }
            }

            return cell;
        }

        /// <summary>
        /// 根据单元格表达式和单元格数据集获取数据
        /// </summary>
        /// <param name="cellExpress">单元格表达式</param>
        /// <param name="workbook">excel工作文件</param>
        /// <param name="currentSheet">当前sheet</param>
        /// <returns></returns>
        public static object GetVByExpress(string cellExpress, IWorkbook workbook, ISheet currentSheet)
        {
            object value = null;

            //含有单元格表达式的取表达式值，没有表达式的取单元格字符串
            if (!string.IsNullOrEmpty(cellExpress) && workbook != null && currentSheet != null)
            {
                IFormulaEvaluator formulaEvaluator = null;
                if (workbook is HSSFWorkbook)
                {
                    formulaEvaluator = new HSSFFormulaEvaluator(workbook);
                }
                else
                {
                    formulaEvaluator = new XSSFFormulaEvaluator(workbook);
                }

                //创建临时行，单元格，执行表达式运算；
                IRow newRow = currentSheet.CreateRow(currentSheet.LastRowNum + 1);
                ICell cell = newRow.CreateCell(0);
                cell.SetCellFormula(cellExpress);
                cell = formulaEvaluator.EvaluateInCell(cell);
                value = cell.ToString();

                currentSheet.RemoveRow(newRow);
            }

            return value ?? "";

        }

        #region 创建工作表

        /// <summary>
        /// 将列表数据生成工作表
        /// </summary>
        /// <param name="tList">要导出的数据集</param>
        /// <param name="fieldNameAndShowNameDic">键值对集合（键：字段名，值：显示名称）</param>
        /// <param name="workbook">更新时添加：要更新的工作表</param>
        /// <param name="sheetName">指定要创建的sheet名称时添加</param>
        /// <param name="excelFileDescription">读取或插入定制需求时添加</param>
        /// <returns></returns>
        public static IWorkbook CreateOrUpdateWorkbook<T>(List<T> tList, Dictionary<string, string> fieldNameAndShowNameDic, IWorkbook workbook = null, string sheetName = "sheet1", ExcelFileDescription excelFileDescription = null) where T : new()
        {
            List<ExcelTitleFieldMapper> titleMapperList = ExcelTitleFieldMapper.GetModelFieldMapper<T>(fieldNameAndShowNameDic);

            workbook = ExcelHelper.CreateOrUpdateWorkbook<T>(tList, titleMapperList, workbook, sheetName, excelFileDescription);
            return workbook;
        }
        /// <summary>
        /// 将列表数据生成工作表（T的属性需要添加：属性名列名映射关系）
        /// </summary>
        /// <param name="tList">要导出的数据集</param>
        /// <param name="workbook">更新时添加：要更新的工作表</param>
        /// <param name="sheetName">指定要创建的sheet名称时添加</param>
        /// <param name="excelFileDescription">读取或插入定制需求时添加</param>
        /// <returns></returns>
        public static IWorkbook CreateOrUpdateWorkbook<T>(List<T> tList, IWorkbook workbook = null, string sheetName = "sheet1", ExcelFileDescription excelFileDescription = null) where T : new()
        {
            List<ExcelTitleFieldMapper> titleMapperList = ExcelTitleFieldMapper.GetModelFieldMapper<T>();

            workbook = ExcelHelper.CreateOrUpdateWorkbook<T>(tList, titleMapperList, workbook, sheetName, excelFileDescription);
            return workbook;
        }

        private static IWorkbook CreateOrUpdateWorkbook<T>(List<T> tList, List<ExcelTitleFieldMapper> titleMapperList, IWorkbook workbook, string sheetName, ExcelFileDescription excelFileDescription = null)
        {
            CellModelColl cellModelColl = new CellModelColl(0);

            int defaultBeginTitleIndex = 0;
            if (excelFileDescription != null)
            {
                defaultBeginTitleIndex = excelFileDescription.TitleRowIndex;
            }

            //补全标题行映射数据的标题和下标位置映射关系
            ISheet sheet = ExcelHelper.GetExcelWorkbookSheet(workbook, sheetName: sheetName);
            IRow titleRow = null;
            if (sheet != null)
            {
                titleRow = sheet.GetRow(defaultBeginTitleIndex);
            }

            if (titleRow != null)
            {
                List<ICell> titleCellList = titleRow.Cells;
                foreach (var titleMapper in titleMapperList)
                {
                    if (titleMapper.ExcelTitleIndex < 0)
                    {
                        foreach (var cellItem in titleCellList)
                        {
                            if (cellItem.ToString().Equals(titleMapper.ExcelTitle, StringComparison.OrdinalIgnoreCase))
                            {
                                titleMapper.ExcelTitleIndex = cellItem.ColumnIndex;
                                break;
                            }
                        }
                    }
                    else if (string.IsNullOrEmpty(titleMapper.ExcelTitle))
                    {
                        ICell cell = titleRow.GetCell(titleMapper.ExcelTitleIndex);
                        if (cell != null)
                        {
                            titleMapper.ExcelTitle = cell.ToString();
                        }
                    }
                }
            }
            else
            {
                //如果是新建Sheet页，则手动初始化下标关系
                for (int i = 0; i < titleMapperList.Count; i++)
                {
                    titleMapperList[i].ExcelTitleIndex = i;
                }
            }

            int currentRowIndex = defaultBeginTitleIndex;
            //添加标题单元格数据
            foreach (var titleMapper in titleMapperList)
            {
                cellModelColl.Add(new CellModel
                {
                    RowIndex = defaultBeginTitleIndex,
                    ColumnIndex = titleMapper.ExcelTitleIndex,
                    CellValue = titleMapper.ExcelTitle,
                    IsCellFormula = false
                });
            }
            currentRowIndex++;

            //将标题行数据转出单元格数据
            foreach (var item in tList)
            {
                foreach (var titleMapper in titleMapperList)
                {
                    cellModelColl.Add(new CellModel
                    {
                        RowIndex = currentRowIndex,
                        ColumnIndex = titleMapper.ExcelTitleIndex,
                        CellValue = titleMapper.PropertyInfo.GetValue(item),
                        IsCellFormula = titleMapper.IsCoordinateExpress
                    });
                }
                currentRowIndex++;
            }

            workbook = ExcelHelper.CreateOrUpdateWorkbook(cellModelColl, workbook, sheetName);

            return workbook;
        }

        /// <summary>
        /// 将单元格数据列表生成工作表
        /// </summary>
        /// <param name="commonCellList">所有的单元格数据列表</param>
        /// <param name="workbook">更新时添加：要更新的工作表</param>
        /// <param name="sheetName">指定要创建的sheet名称时添加</param>
        /// <returns></returns>
        public static IWorkbook CreateOrUpdateWorkbook(CellModelColl commonCellList, IWorkbook workbook = null, string sheetName = "sheet1")
        {
            //xls文件格式属于老版本文件，一个sheet最多保存65536行；而xlsx属于新版文件类型；
            //Excel 07 - 2003一个工作表最多可有65536行，行用数字1—65536表示; 最多可有256列，列用英文字母A—Z，AA—AZ，BA—BZ，……，IA—IV表示；一个工作簿中最多含有255个工作表，默认情况下是三个工作表；
            //Excel 2007及以后版本，一个工作表最多可有1048576行，16384列；
            if (workbook == null)
            {
                workbook = new XSSFWorkbook();
                //workbook = new HSSFWorkbook();
            }
            ISheet worksheet = ExcelHelper.GetExcelWorkbookSheet(workbook, sheetName);

            if (worksheet != null && commonCellList != null && commonCellList.Count > 0)
            {
                //设置首列显示
                IRow row1 = null;
                int rowIndex = 0;
                int maxRowIndex = commonCellList.Max(m => m.RowIndex);
                Dictionary<int, CellModel> rowColumnIndexCellDIC = null;
                ICell cell = null;
                object cellValue = null;

                do
                {
                    rowColumnIndexCellDIC = commonCellList.GetRawCellList(rowIndex).ToDictionary(m => m.ColumnIndex);
                    int maxColumnIndex = rowColumnIndexCellDIC.Count > 0 ? rowColumnIndexCellDIC.Keys.Max() : 0;

                    if (rowColumnIndexCellDIC != null && rowColumnIndexCellDIC.Count > 0)
                    {
                        row1 = worksheet.GetRow(rowIndex);
                        if (row1 == null)
                        {
                            row1 = worksheet.CreateRow(rowIndex);
                        }
                        int columnIndex = 0;
                        do
                        {
                            cell = row1.GetCell(columnIndex);
                            if (cell == null)
                            {
                                cell = row1.CreateCell(columnIndex);
                            }

                            if (rowColumnIndexCellDIC.ContainsKey(columnIndex))
                            {
                                cellValue = rowColumnIndexCellDIC[columnIndex].CellValue;

                                CellFactory.SetCellValue(cell, cellValue, outputFormat: null, rowColumnIndexCellDIC[columnIndex].IsCellFormula);
                            }
                            columnIndex++;
                        } while (columnIndex <= maxColumnIndex);
                    }
                    rowIndex++;
                } while (rowIndex <= maxRowIndex);

                //设置表达式重算（如果不添加该代码，表达式更新不出结果值）
                worksheet.ForceFormulaRecalculation = true;
            }

            return workbook;
        }

        /// <summary>
        /// 更新模板文件数据：将使用单元格映射的数据T存入模板文件中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="t"></param>
        /// <param name="excelFileDescription"></param>
        /// <returns></returns>
        public static IWorkbook UpdateTemplateWorkbook<T>(IWorkbook workbook, ISheet sheet, T t, ExcelFileDescription excelFileDescription = null)
        {
            //该方法默认替换模板数据在首个sheet里

            CellModelColl commonCellColl = ExcelHelper.ReadCellList(workbook, sheet, false);

            List<IExcelCellPointDeepUpdate> excelCellPointDeepList = new List<IExcelCellPointDeepUpdate>(0);
            if (excelFileDescription != null)
            {
                excelCellPointDeepList.Add((IExcelCellPointDeepUpdate)excelFileDescription.ExcelDeepUpdateList);
            }

            //获取t的单元格映射列表
            List<ExcelCellFieldMapper> cellMapperList = ExcelCellFieldMapper.GetModelFieldMapper<T>();
            foreach (var cellMapper in cellMapperList)
            {
                if (cellMapper.CellParamWriteList.Count > 0)
                {
                    foreach (var cellParamWriteAttribute in cellMapper.CellParamWriteList)
                    {
                        CellModel cellModel = commonCellColl.GetCell(cellParamWriteAttribute.CellParamName);
                        if (cellModel != null)
                        {
                            cellModel.CellValue = cellMapper.PropertyInfo.GetValue(t);
                        }
                    }
                }
                if (cellMapper.CellPointWriteList.Count > 0)
                {
                    object cellValue = cellMapper.PropertyInfo.GetValue(t);
                    ICellModel firstCellPosition = null;
                    foreach (var cellPointWriteAttribute in cellMapper.CellPointWriteList)
                    {
                        firstCellPosition = CellFactory.GetCellByExcelPosition(cellPointWriteAttribute.CellPosition);
                        CellFactory.SetDeepUpdateCellValue(sheet, firstCellPosition.RowIndex, firstCellPosition.ColumnIndex, cellValue, cellPointWriteAttribute.OutputFormat, false, excelCellPointDeepList);
                    }
                }
            }

            workbook = ExcelHelper.CreateOrUpdateWorkbook(commonCellColl, workbook, sheet.SheetName);

            return workbook;
        }

        #endregion

        #region 保存工作表到文件

        /// <summary>
        /// 保存Workbook数据为文件
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fileDirectoryPath"></param>
        /// <param name="fileName"></param>
        public static void SaveWorkbookToFile(IWorkbook workbook, string filePath)
        {
            //xls文件格式属于老版本文件，一个sheet最多保存65536行；而xlsx属于新版文件类型；
            //Excel 07 - 2003一个工作表最多可有65536行，行用数字1—65536表示; 最多可有256列，列用英文字母A—Z，AA—AZ，BA—BZ，……，IA—IV表示；一个工作簿中最多含有255个工作表，默认情况下是三个工作表；
            //Excel 2007及以后版本，一个工作表最多可有1048576行，16384列；

            MemoryStream ms = new MemoryStream();
            //这句代码非常重要，如果不加，会报：打开的EXCEL格式与扩展名指定的格式不一致
            ms.Seek(0, SeekOrigin.Begin);
            workbook.Write(ms);
            byte[] myByteArray = ms.GetBuffer();

            string fileDirectoryPath = filePath.Split('\\')[0];
            if (!Directory.Exists(fileDirectoryPath))
            {
                Directory.CreateDirectory(fileDirectoryPath);
            }
            string fileName = filePath.Replace(fileDirectoryPath, "");

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            File.WriteAllBytes(filePath, myByteArray);
        }

        /// <summary>
        /// 保存workbook到字节流中(提供给API接口使用)
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static byte[] SaveWorkbookToByte(IWorkbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            stream.Seek(0, SeekOrigin.Begin);
            workbook.Write(stream);

            byte[] byteArray = stream.GetBuffer();
            return byteArray;
        }

        #endregion

        #region 读取Excel数据

        /// <summary>
        /// 读取Excel数据1_手动提供属性信息和标题对应关系
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="fieldNameAndShowNameDic"></param>
        /// <param name="excelFileDescription"></param>
        /// <returns></returns>
        public static List<T> ReadTitleDataList<T>(string filePath, Dictionary<string, string> fieldNameAndShowNameDic, ExcelFileDescription excelFileDescription) where T : new()
        {
            //标题属性字典列表
            List<ExcelTitleFieldMapper> titleMapperList = ExcelTitleFieldMapper.GetModelFieldMapper<T>(fieldNameAndShowNameDic);

            List<T> tList = ExcelHelper._GetTList<T>(filePath, titleMapperList, excelFileDescription);
            return tList ?? new List<T>(0);
        }

        /// <summary>
        /// 读取Excel数据2_使用Excel标记特性和文件描述自动创建关系
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="excelFileDescription"></param>
        /// <returns></returns>
        public static List<T> ReadTitleDataList<T>(string filePath, ExcelFileDescription excelFileDescription) where T : new()
        {
            //标题属性字典列表
            List<ExcelTitleFieldMapper> titleMapperList = ExcelTitleFieldMapper.GetModelFieldMapper<T>();

            List<T> tList = ExcelHelper._GetTList<T>(filePath, titleMapperList, excelFileDescription);
            return tList ?? new List<T>(0);
        }

        private static List<T> _GetTList<T>(string filePath, List<ExcelTitleFieldMapper> titleMapperList, ExcelFileDescription excelFileDescription) where T : new()
        {
            List<T> tList = new List<T>(500 * 10000);
            T t = default(T);

            try
            {
                IWorkbook workbook = ExcelHelper.GetExcelWorkbook(filePath);
                IFormulaEvaluator formulaEvaluator = null;

                if (workbook is XSSFWorkbook)
                {
                    formulaEvaluator = new XSSFFormulaEvaluator(workbook);
                }
                else if (workbook is HSSFWorkbook)
                {
                    formulaEvaluator = new HSSFFormulaEvaluator(workbook);
                }

                int sheetCount = workbook.NumberOfSheets;

                int currentSheetIndex = 0;
                int currentSheetRowTitleIndex = -1;
                do
                {
                    var sheet = workbook.GetSheetAt(currentSheetIndex);

                    //标题下标属性字典
                    Dictionary<int, ExcelTitleFieldMapper> sheetTitleIndexPropertyDic = new Dictionary<int, ExcelTitleFieldMapper>(0);

                    //如果没有设置标题行，则通过自动查找方法获取
                    if (excelFileDescription.TitleRowIndex < 0)
                    {
                        string[] titleArray = titleMapperList.Select(m => m.ExcelTitle).ToArray();
                        currentSheetRowTitleIndex = ExcelHelper.GetSheetTitleIndex(sheet, titleArray);
                    }
                    else
                    {
                        currentSheetRowTitleIndex = excelFileDescription.TitleRowIndex;
                    }

                    var rows = sheet.GetRowEnumerator();

                    bool isHaveTitleIndex = false;
                    //含有Excel行下标
                    if (titleMapperList.Count > 0 && titleMapperList[0].ExcelTitleIndex >= 0)
                    {
                        isHaveTitleIndex = true;

                        foreach (var titleMapper in titleMapperList)
                        {
                            sheetTitleIndexPropertyDic.Add(titleMapper.ExcelTitleIndex, titleMapper);
                        }
                    }

                    PropertyInfo propertyInfo = null;
                    int currentRowIndex = 0;

                    if (currentSheetRowTitleIndex >= 0)
                    {
                        while (rows.MoveNext())
                        {
                            IRow row = (IRow)rows.Current;
                            currentRowIndex = row.RowNum;

                            //到达标题行（寻找标题行映射）
                            if (isHaveTitleIndex == false && currentRowIndex == currentSheetRowTitleIndex)
                            {
                                ICell cell = null;
                                string cellValue = null;
                                Dictionary<string, ExcelTitleFieldMapper> titleMapperDic = titleMapperList.ToDictionary(m => m.ExcelTitle);
                                for (int i = 0; i < row.Cells.Count; i++)
                                {
                                    cell = row.Cells[i];
                                    cellValue = cell.StringCellValue;
                                    if (titleMapperDic.ContainsKey(cellValue))
                                    {
                                        sheetTitleIndexPropertyDic.Add(i, titleMapperDic[cellValue]);
                                    }
                                }
                            }

                            //到达内容行
                            if (currentRowIndex > currentSheetRowTitleIndex)
                            {
                                t = new T();
                                ExcelTitleFieldMapper excelTitleFieldMapper = null;
                                foreach (var titleIndexItem in sheetTitleIndexPropertyDic)
                                {
                                    ICell cell = row.GetCell(titleIndexItem.Key);

                                    excelTitleFieldMapper = titleIndexItem.Value;

                                    //没有数据的单元格默认为null
                                    string cellValue = cell?.ToString() ?? "";
                                    propertyInfo = excelTitleFieldMapper.PropertyInfo;
                                    try
                                    {
                                        if (excelTitleFieldMapper.IsCheckContentEmpty)
                                        {
                                            if (string.IsNullOrEmpty(cellValue))
                                            {
                                                t = default(T);
                                                break;
                                            }
                                        }

                                        if (excelTitleFieldMapper.IsCoordinateExpress || (cell != null && cell.CellType == CellType.Formula))
                                        {
                                            //读取含有表达式的单元格值
                                            cellValue = formulaEvaluator.Evaluate(cell).StringValue;
                                            propertyInfo.SetValue(t, Convert.ChangeType(cellValue, propertyInfo.PropertyType));
                                        }
                                        else if (propertyInfo.PropertyType.IsEnum)
                                        {
                                            object enumObj = propertyInfo.PropertyType.InvokeMember(cellValue, BindingFlags.GetField, null, null, null);
                                            propertyInfo.SetValue(t, Convert.ChangeType(enumObj, propertyInfo.PropertyType));
                                        }
                                        else
                                        {
                                            propertyInfo.SetValue(t, Convert.ChangeType(cellValue, propertyInfo.PropertyType));
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        ExcelHelper._Logger.Debug($"文件_{filePath}读取{currentRowIndex + 1}行内容失败！");
                                        t = default(T);
                                        break;
                                    }
                                }
                                if (t != null)
                                {
                                    tList.Add(t);
                                }
                            }
                        }
                    }

                    currentSheetIndex++;

                } while (currentSheetIndex + 1 <= sheetCount);
            }
            catch (Exception e)
            {
                throw new Exception($"文件：{filePath}被占用！", e);
            }
            return tList ?? new List<T>(0);
        }

        private static List<T> _GetTListBySheet<T>(ISheet sheet, ExcelFileDescription excelFileDescription) where T : new()
        {
            List<ExcelTitleFieldMapper> titleMapperList = ExcelTitleFieldMapper.GetModelFieldMapper<T>();

            List<T> tList = new List<T>(500 * 10000);
            T t = default(T);

            IWorkbook workbook = sheet.Workbook;
            IFormulaEvaluator formulaEvaluator = null;

            if (workbook is XSSFWorkbook)
            {
                formulaEvaluator = new XSSFFormulaEvaluator(workbook);
            }
            else if (workbook is HSSFWorkbook)
            {
                formulaEvaluator = new HSSFFormulaEvaluator(workbook);
            }

            //标题下标属性字典
            Dictionary<int, ExcelTitleFieldMapper> sheetTitleIndexPropertyDic = new Dictionary<int, ExcelTitleFieldMapper>(0);

            //如果没有设置标题行，则通过自动查找方法获取
            int currentSheetRowTitleIndex = 0;
            if (excelFileDescription.TitleRowIndex < 0)
            {
                string[] titleArray = titleMapperList.Select(m => m.ExcelTitle).ToArray();
                currentSheetRowTitleIndex = ExcelHelper.GetSheetTitleIndex(sheet, titleArray);
            }
            else
            {
                currentSheetRowTitleIndex = excelFileDescription.TitleRowIndex;
            }

            var rows = sheet.GetRowEnumerator();

            bool isHaveTitleIndex = false;
            //含有Excel行下标
            if (titleMapperList.Count > 0 && titleMapperList[0].ExcelTitleIndex >= 0)
            {
                isHaveTitleIndex = true;

                foreach (var titleMapper in titleMapperList)
                {
                    sheetTitleIndexPropertyDic.Add(titleMapper.ExcelTitleIndex, titleMapper);
                }
            }

            PropertyInfo propertyInfo = null;
            int currentRowIndex = 0;

            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                currentRowIndex = row.RowNum;

                //到达标题行
                if (isHaveTitleIndex == false && currentRowIndex == currentSheetRowTitleIndex)
                {
                    ICell cell = null;
                    string cellValue = null;
                    Dictionary<string, ExcelTitleFieldMapper> titleMapperDic = titleMapperList.ToDictionary(m => m.ExcelTitle);
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        cell = row.Cells[i];
                        cellValue = cell.StringCellValue;
                        if (titleMapperDic.ContainsKey(cellValue))
                        {
                            sheetTitleIndexPropertyDic.Add(i, titleMapperDic[cellValue]);
                        }
                    }
                }

                //到达内容行
                if (currentRowIndex > currentSheetRowTitleIndex)
                {
                    t = new T();
                    ExcelTitleFieldMapper excelTitleFieldMapper = null;
                    foreach (var titleIndexItem in sheetTitleIndexPropertyDic)
                    {
                        ICell cell = row.GetCell(titleIndexItem.Key);

                        excelTitleFieldMapper = titleIndexItem.Value;

                        //没有数据的单元格默认为null
                        string cellValue = cell?.ToString() ?? "";
                        propertyInfo = excelTitleFieldMapper.PropertyInfo;
                        try
                        {
                            if (excelTitleFieldMapper.IsCheckContentEmpty)
                            {
                                if (string.IsNullOrEmpty(cellValue))
                                {
                                    t = default(T);
                                    break;
                                }
                            }

                            if (excelTitleFieldMapper.IsCoordinateExpress || cell.CellType == CellType.Formula)
                            {
                                //读取含有表达式的单元格值
                                cellValue = formulaEvaluator.Evaluate(cell).StringValue;
                                propertyInfo.SetValue(t, Convert.ChangeType(cellValue, propertyInfo.PropertyType));
                            }
                            else if (propertyInfo.PropertyType.IsEnum)
                            {
                                object enumObj = propertyInfo.PropertyType.InvokeMember(cellValue, BindingFlags.GetField, null, null, null);
                                propertyInfo.SetValue(t, Convert.ChangeType(enumObj, propertyInfo.PropertyType));
                            }
                            else
                            {
                                propertyInfo.SetValue(t, Convert.ChangeType(cellValue, propertyInfo.PropertyType));
                            }
                        }
                        catch (Exception e)
                        {
                            ExcelHelper._Logger.Debug($"sheetName_{sheet.SheetName}读取{currentRowIndex + 1}行内容失败！");
                            t = default(T);
                            break;
                        }
                    }
                    if (t != null)
                    {
                        tList.Add(t);
                    }
                }
            }

            return tList ?? new List<T>(0);
        }

        public static CellModelColl ReadCellList(IWorkbook workbook, ISheet sheet, bool isRunFormula = false)
        {
            CellModelColl commonCells = new CellModelColl(10000);

            IFormulaEvaluator formulaEvaluator = null;
            if (workbook != null)
            {
                if (workbook is HSSFWorkbook)
                {
                    formulaEvaluator = new HSSFFormulaEvaluator(workbook);
                }
                else
                {
                    formulaEvaluator = new XSSFFormulaEvaluator(workbook);
                }
            }
            if (sheet != null)
            {
                CellModel cellModel = null;

                var rows = sheet.GetRowEnumerator();

                //从第1行数据开始获取
                while (rows.MoveNext())
                {
                    IRow row = (IRow)rows.Current;

                    List<ICell> cellList = row.Cells;

                    ICell cell = null;
                    foreach (var cellItem in cellList)
                    {
                        cell = cellItem;
                        if (isRunFormula && cell.CellType == CellType.Formula)
                        {
                            cell = formulaEvaluator.EvaluateInCell(cell);
                        }

                        cellModel = new CellModel
                        {
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex,
                            CellValue = cell.ToString(),
                            IsCellFormula = cell.CellType == CellType.Formula
                        };

                        commonCells.Add(cellModel);
                    }
                }
            }
            return commonCells;
        }

        /// <summary>
        /// 获取文件单元格数据对象
        /// </summary>
        /// <typeparam name="T">T的属性必须标记了ExcelCellAttribute</typeparam>
        /// <param name="filePath">文建路径</param>
        /// <param name="sheetIndex">（可选）sheet所在位置</param>
        /// <param name="sheetName">（可选）sheet名称</param>
        /// <returns></returns>
        public static T ReadCellData<T>(IWorkbook workbook, ISheet sheet) where T : new()
        {
            T t = new T();

            if (workbook != null)
            {

                if (sheet != null)
                {
                    Dictionary<PropertyInfo, ExcelCellFieldMapper> propertyMapperDic = ExcelCellFieldMapper.GetModelFieldMapper<T>().ToDictionary(m => m.PropertyInfo);
                    string cellExpress = null;
                    string pValue = null;
                    PropertyInfo propertyInfo = null;
                    foreach (var item in propertyMapperDic)
                    {
                        cellExpress = item.Value.CellExpressRead.CellCoordinateExpress;
                        propertyInfo = item.Key;
                        pValue = ExcelHelper.GetVByExpress(cellExpress, workbook, sheet).ToString();
                        if (!string.IsNullOrEmpty(pValue))
                        {
                            try
                            {
                                propertyInfo.SetValue(t, Convert.ChangeType(pValue, propertyInfo.PropertyType));
                            }
                            catch (Exception)
                            {

                                throw;
                            }

                        }
                    }
                }
            }

            return t;
        }

        /// <summary>
        /// 读取单元格数据对象列表-支持深度读取
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="excelFileDescription"></param>
        /// <returns></returns>
        public static List<T> ReadCellData<T>(IWorkbook workbook, ISheet sheet, ExcelFileDescription excelFileDescription) where T : new()
        {
            List<T> tList = new List<T>(0);
            T t = default(T);

            #region 获取深度表达式更新列表

            List<IExcelCellExpressDeepUpdate<T>> excelCellExpressDeepUpdateList = new List<IExcelCellExpressDeepUpdate<T>>(0);
            if (excelFileDescription != null)
            {
                foreach (var item in excelFileDescription.ExcelDeepUpdateList)
                {
                    if (item is IExcelCellExpressDeepUpdate<T>)
                    {
                        excelCellExpressDeepUpdateList.Add((IExcelCellExpressDeepUpdate<T>)item);
                    }
                }
            }

            #endregion

            #region 通过表达式映射列表读取对象T

            Func<List<ExcelCellFieldMapper>, T> expressMapperFunc = (excelCellFieldMapperList) =>
            {
                t = new T();
                foreach (var cellMapper in excelCellFieldMapperList)
                {
                    string currentCellExpress = cellMapper.CellExpressRead.CellCoordinateExpress;

                    object pValue = ExcelHelper.GetVByExpress(currentCellExpress, workbook, sheet);

                    try
                    {
                        cellMapper.PropertyInfo.SetValue(t, Convert.ChangeType(pValue, cellMapper.PropertyInfo.PropertyType));
                    }
                    catch (Exception)
                    {
                    }
                }
                return t;
            };

            #endregion

            #region 执行初始表达式数据收集

            //获取t的单元格映射列表
            List<ExcelCellFieldMapper> cellMapperList = ExcelCellFieldMapper.GetModelFieldMapper<T>();
            t = expressMapperFunc(cellMapperList);

            #endregion

            #region 执行深度更新策略收集数据

            Action<IExcelCellExpressDeepUpdate<T>> actionDeepReadAction = (excelCellExpressDeepUpdate) =>
            {
                //获取初始表达式映射列表
                cellMapperList = ExcelCellFieldMapper.GetModelFieldMapper<T>();

                //执行单元格表达式深度更新

                bool isContinute = false;

                do
                {
                    //通过深度更新策略更新初始表达式数据
                    foreach (var cellMapper in cellMapperList)
                    {
                        if (cellMapper.CellExpressRead != null)
                        {
                            string currentCellExpress = cellMapper.CellExpressRead.CellCoordinateExpress;
                            currentCellExpress = excelCellExpressDeepUpdate.GetNextCellExpress(currentCellExpress);
                            cellMapper.CellExpressRead.CellCoordinateExpress = currentCellExpress;
                        }
                    }
                    t = expressMapperFunc(cellMapperList);
                    isContinute = excelCellExpressDeepUpdate.IsContinute(t);
                    if (isContinute)
                    {
                        tList.Add(t);
                    }

                } while (isContinute);
            };

            foreach (var item in excelCellExpressDeepUpdateList)
            {
                actionDeepReadAction(item);
            }

            #endregion

            return tList;
        }

        /// <summary>
        /// 获取文件首个sheet的标题位置
        /// </summary>
        /// <typeparam name="T">T必须做了标题映射</typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static int FileFirstSheetTitleIndex<T>(string filePath)
        {
            int titleIndex = 0;

            if (File.Exists(filePath))
            {
                try
                {
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook = null;
                        try
                        {
                            workbook = new XSSFWorkbook(fileStream);
                        }
                        catch (Exception)
                        {
                            workbook = new HSSFWorkbook(fileStream);
                        }

                        string[] titleArray = ExcelTitleFieldMapper.GetModelFieldMapper<T>().Select(m => m.ExcelTitle).ToArray();

                        ISheet sheet = workbook.GetSheetAt(0);
                        titleIndex = ExcelHelper.GetSheetTitleIndex(sheet, titleArray);
                    }
                }
                catch (Exception e)
                {
                    throw new Exception($"文件：{filePath}被占用！", e);
                }
            }

            return titleIndex;
        }

        /// <summary>
        /// 获取文件首个sheet的标题位置
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="titleNames"></param>
        /// <returns></returns>
        public static int FileFirstSheetTitleIndex(string filePath, params string[] titleNames)
        {
            int titleIndex = 0;

            if (File.Exists(filePath))
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = null;
                    try
                    {
                        workbook = new XSSFWorkbook(fileStream);
                    }
                    catch (Exception)
                    {
                        workbook = new HSSFWorkbook(fileStream);
                    }
                    ISheet sheet = workbook.GetSheetAt(0);
                    titleIndex = ExcelHelper.GetSheetTitleIndex(sheet, titleNames);
                }
            }

            return titleIndex;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 根据标题名称获取标题行下标位置
        /// </summary>
        /// <param name="sheet">要查找的sheet</param>
        /// <param name="titleNames">标题名称</param>
        /// <returns></returns>
        private static int GetSheetTitleIndex(ISheet sheet, params string[] titleNames)
        {
            int titleIndex = -1;

            if (sheet != null && titleNames != null && titleNames.Length > 0)
            {
                var rows = sheet.GetRowEnumerator();
                List<ICell> cellList = null;
                List<string> rowValueList = null;

                //从第1行数据开始获取
                while (rows.MoveNext())
                {
                    IRow row = (IRow)rows.Current;

                    cellList = row.Cells;
                    rowValueList = new List<string>(cellList.Count);
                    foreach (var cell in cellList)
                    {
                        rowValueList.Add(cell.ToString());
                    }

                    bool isTitle = true;
                    foreach (var title in titleNames)
                    {
                        if (!rowValueList.Contains(title))
                        {
                            isTitle = false;
                            break;
                        }
                    }
                    if (isTitle)
                    {
                        titleIndex = row.RowNum;
                        break;
                    }
                }
            }
            return titleIndex;
        }

        #endregion

    }

}
