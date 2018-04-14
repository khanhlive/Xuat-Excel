using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Web;
using XuatExcelClosedXML.ExcelCore;

namespace XuatExcelClosedXML.Design
{
    public class ExcelCommon
    {
        public static List<ExcelColumnContent> GetColumnsValid(DataTable data, List<ExcelColumnContent> columns)
        {
            List<ExcelColumnContent> list = new List<ExcelColumnContent>();
            foreach (var item in columns)
            {
                if (data.Columns.Contains(item.Name))
                {
                    list.Add(item);
                }
            }
            return list;
         }

        public static T DeepCopy<T>(T source)
        {
            var serialized = JsonConvert.SerializeObject(source);
            return JsonConvert.DeserializeObject<T>(serialized);
        }
    }
}