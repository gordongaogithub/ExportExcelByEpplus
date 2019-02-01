using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelByEpplus
{
    public static class EpplusHelper
    {
        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headerTexts"></param>
        public static void AddHeader(ExcelWorksheet sheet, params string[] headerTexts)
        {
            for (var i = 0; i < headerTexts.Length; i++)
            {
                AddHeader(sheet, i + 1, headerTexts[i]);
            }
        }

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="headerText"></param>
        public static void AddHeader(ExcelWorksheet sheet, int columnIndex, string headerText)
        {
            sheet.Cells[1, columnIndex].Value = headerText;
            sheet.Cells[1, columnIndex].Style.Font.Bold = true;
        }

        /// <summary>
        /// 添加数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="items"></param>
        /// <param name="propertySelectors"></param>
        public  static void AddObjects(ExcelWorksheet sheet, int startRowIndex, IList<Student> items, Func<Student, object>[] propertySelectors)
        {
            for (var i = 0; i < items.Count; i++)
            {
                for (var j = 0; j < propertySelectors.Length; j++)
                {
                    sheet.Cells[i + startRowIndex, j + 1].Value = propertySelectors[j](items[i]);
                }
            }
        }
    }
}
