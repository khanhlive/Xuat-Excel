using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XuatExcelClosedXML.ExcelCore
{
    public class ExcelColumnContent
    {
        public string Name { get; set; }
        public double Width { get; set; }
        public string Caption { get; set; }
        public XLAlignmentHorizontalValues HorizontalAlignment { get; set; }
        public ExcelColumnContent()
        {
            this.HorizontalAlignment = XLAlignmentHorizontalValues.Left;
        }

    }
}