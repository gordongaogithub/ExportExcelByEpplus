using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelByEpplus
{
    class Program
    {
        static void Main(string[] args)
        {
            //获得数据
            List<Student> studentList = new List<Student>();
            for (int i = 0; i < 10; i++)
            {
                Student s = new Student();
                s.Code = "c" + i;
                s.Name = "s" + i;
                studentList.Add(s);
            }

            //创建excel
            string fileName = @"d:\" + "导出excel" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            FileInfo newFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                List<ExcelExportDto<Student>> excelExportDtoList = new List<ExcelExportDto<Student>>();
                excelExportDtoList.Add(new ExcelExportDto<Student>("Code", _ => _.Code));
                excelExportDtoList.Add(new ExcelExportDto<Student>("Name", _ => _.Name));

                List<string> columnsNameList = new List<string>();
                List<Func<Student, object>> columnsValueList = new List<Func<Student, object>>();
                foreach (var item in excelExportDtoList)
                {
                    columnsNameList.Add(item.ColumnName);
                    columnsValueList.Add(item.ColumnValue);
                }

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Test");
                worksheet.OutLineApplyStyle = true;
                //添加表头
                EpplusHelper.AddHeader(worksheet, columnsNameList.ToArray());
                //添加数据
                EpplusHelper.AddObjects(worksheet, 2, studentList, columnsValueList.ToArray());
                package.Save();
            }
        }
    }
}
