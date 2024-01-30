using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace PPPayReportTools.Excel
{
    /// <summary>
    /// 单元格字段映射类
    /// </summary>
    internal class ExcelCellFieldMapper
    {
        /// <summary>
        /// 属性信息(一个属性可以添加一个表达式读取，多个变量替换和多个坐标写入)
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 单元格—表达式读取（单元格坐标表达式（如：A1,B2,C1+C2...横坐标使用26进制字母，纵坐标使用十进制数字））
        /// </summary>
        public ExcelCellExpressReadAttribute CellExpressRead { get; set; }

        /// <summary>
        /// 单元格—模板文件的预定义变量写入（{A} {B}）
        /// </summary>
        public List<ExcelCellParamWriteAttribute> CellParamWriteList { get; set; }

        /// <summary>
        /// 单元格—坐标位置写入（(0,0),(1,1)）
        /// </summary>
        public List<ExcelCellPointWriteAttribute> CellPointWriteList { get; set; }

        /// <summary>
        /// 获取对应关系_T属性添加了单元格映射关系
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static List<ExcelCellFieldMapper> GetModelFieldMapper<T>()
        {
            List<ExcelCellFieldMapper> fieldMapperList = new List<ExcelCellFieldMapper>(100);

            List<PropertyInfo> tPropertyInfoList = typeof(T).GetProperties().ToList();
            ExcelCellExpressReadAttribute cellExpress = null;
            List<ExcelCellParamWriteAttribute> cellParamWriteList = null;
            List<ExcelCellPointWriteAttribute> cellPointWriteList = null;
            foreach (var item in tPropertyInfoList)
            {
                cellExpress = item.GetCustomAttribute<ExcelCellExpressReadAttribute>();
                cellParamWriteList = item.GetCustomAttributes<ExcelCellParamWriteAttribute>().ToList();
                cellPointWriteList = item.GetCustomAttributes<ExcelCellPointWriteAttribute>().ToList();
                if (cellExpress != null || cellParamWriteList.Count > 0 || cellPointWriteList.Count > 0)
                {
                    fieldMapperList.Add(new ExcelCellFieldMapper
                    {
                        CellExpressRead = cellExpress,
                        CellParamWriteList = cellParamWriteList,
                        CellPointWriteList = cellPointWriteList,
                        PropertyInfo = item
                    });
                }
            }

            return fieldMapperList;
        }
    }
}
