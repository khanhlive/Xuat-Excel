using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class ExportExcelHelper
    {
        public static int GetIndexRowFromString(string cell)
        {
            char[] chars=cell.ToCharArray();
            bool flag = false;
            string str = "";
            for (int i = 0; i < chars.Length; i++)
            {
                if (char.IsNumber(chars[i]))
                {
                    flag = true;
                }
                if (flag) str += chars[i].ToString();
            }
            return Convert.ToInt32(str);
        }
    }
}