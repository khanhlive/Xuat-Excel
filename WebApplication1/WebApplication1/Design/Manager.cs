using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using WebApplication1.ExcelCore;

namespace WebApplication1.Design
{
    public class BookManager
    {
        private string _filename;
        List<WorksheetTemplate> sheetTemplate;
        ClosedXML.Excel.XLWorkbook workbook;

        public BookManager(string filename)
        {
            this.sheetTemplate = new List<WorksheetTemplate>();
            this._filename = filename;
        }

        public void Export()
        {
            using (this.workbook=new ClosedXML.Excel.XLWorkbook())
            {
                foreach (var item in sheetTemplate)
                {
                    var sheet = this.workbook.AddWorksheet(item.Name);
                    sheet = item.GetTemplate(sheet);
                }
                this.Response();
            }
        }
        
        protected virtual void Response()
        {
            HttpResponse _response = HttpContext.Current.Response;
            _response.ClearContent();
            _response.Buffer = true;
            _response.AddHeader("content-disposition", string.Format("attachment; filename={0}.xlsx", this._filename ?? (new Guid()).ToString()));
            _response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            _response.Charset = "";
            using (MemoryStream MyMemoryStream = new MemoryStream())
            {
                workbook.SaveAs(MyMemoryStream);
                MyMemoryStream.WriteTo(_response.OutputStream);
                _response.Flush();
                _response.End();
            }
        }

        public void AddSheet(WorksheetTemplate sheet)
        {
            this.sheetTemplate.Add(sheet);
        }
    }
    public class SheetManager
    {
    }
}