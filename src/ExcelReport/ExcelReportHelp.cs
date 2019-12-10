using ExcelReport;
using ExcelReport.Model;
using ExcelReport.Renderers;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReport
{
    public class ExcelReportHelp
    {
        /// <summary>
        /// 导出Excel Form 
        /// </summary>
        /// <param name="fileTemplate"></param>
        /// <param name="pairs"></param>
        /// <returns></returns>
        public static Response<byte[]> OutPutExcelBytes(string fileTemplate, params (string, object)[] pairs)
        {
            return OutPutExcelBytes(fileTemplate, "Sheet1", pairs);
        }

        public static Response<byte[]> OutPutExcelBytes(string fileTemplate, string sheelName, params (string, object)[] pairs)
        {

            Response<byte[]> response = new Response<byte[]>(true);
            if (!System.IO.File.Exists(fileTemplate))
            {
                response.Success = false; response.MessageText = "文件不存在"; return response;
            }
            //1、构建数据结构
            List<IElementRenderer> list = new List<IElementRenderer>();
            FactoryRenderer(pairs, list);

            //2、填充
            var render = new SheetRenderer(sheelName, list.ToArray());
            //3、渲染输出
            response.Items = Export.ExportToBuffer(fileTemplate, render);
            ExportHelper.ExportToLocal(fileTemplate, "out.xls", render);
            return response;
        }

        private static void FactoryRenderer((string, object)[] pairs, List<IElementRenderer> list)
        {
            foreach (var item in pairs)
            {
                if (item.Item2.GetType().IsGenericType)
                {
                    var chlist = (IEnumerable<object>)item.Item2;
                    if (chlist.Count() > 0)
                    {
                        list.Add(FactoryRepeaterRenderer<object>(item.Item1, chlist));
                    }
                }
                else
                {
                    list.Add(new ParameterRenderer(item.Item1, item.Item2));
                }
            }
        }

        private static RepeaterRenderer<T> FactoryRepeaterRenderer<T>(string Key, IEnumerable<T> data) where T : class
        {
            RepeaterRenderer<T> repeater = new RepeaterRenderer<T>(Key, data);
            var propers = data.First().GetType().GetProperties();
            var num = 1;
            foreach (var item in propers)
            {
                repeater.Append(new ParameterRenderer<T>(item.Name, a => item.GetValue(a)));
            }
            repeater.Append(new ParameterRenderer<T>("No", a => num++));
            return repeater;
        }
    }

}
