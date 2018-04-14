using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using WebApplication1.ExcelCore.Interface;

namespace WebApplication1.ExcelCore
{
    public class WorkbookManager : IExportExcel
    {
        private string _filename;
        List<WorkSheetSetting> sheets;
        ClosedXML.Excel.XLWorkbook workbook;

        public WorkbookManager(string filename)
        {
            this.sheets = new List<WorkSheetSetting>();
            this._filename = filename;
        }

        public void Export()
        {
            using (this.workbook=new ClosedXML.Excel.XLWorkbook())
            {
                foreach (var item in sheets)
                {
                    var sheet = this.workbook.AddWorksheet(item.SheetName);
                    WorksheetManager sh = new WorksheetManager(sheet);
                    sh.Setting = item;
                    sheet = sh.RenderSheet();
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

        public void AddSheet(WorkSheetSetting sheetSetting)
        {
            this.sheets.Add(sheetSetting);
        }
    }

    public class WorkSheetSetting :ICloneable
    {
        protected XLPaperSize paperSize = XLPaperSize.A4Paper;
        protected bool tableHeaderBold;
        protected XLPageOrientation pageOrientation = XLPageOrientation.Default;
        protected string titleSheet;
        //protected int row;
        //protected IXLRange tableRange;
        protected XLBorderStyleValues boderStyle;
        protected bool _wrapText = true;
        protected string sheetName;
        protected bool _showGridLine = false;
        //protected IXLWorksheet worksheet;
        protected DataTable dataSource;
        //protected List<IXLCell> headerCells = new List<IXLCell>();
        protected ExcelColumnContent[] columnsWidth;
        protected XLAlignmentHorizontalValues headerHorizontalAlignment;

        public XLPageOrientation PageOrientation
        {
            get { return pageOrientation; }
            set { pageOrientation = value; }
        }

        public XLPaperSize PaperSize
        {
            get { return paperSize; }
            set { paperSize = value; }
        }

        public bool TableHeaderBold
        {
            get { return tableHeaderBold; }
            set { tableHeaderBold = value; }
        }

        public ExcelColumnContent[] ColumnsWidth
        {
            get { return columnsWidth; }
            set { columnsWidth = value; }
        }

        public XLAlignmentHorizontalValues HeaderHorizontalAlignment
        {
            get { return headerHorizontalAlignment; }
            set { headerHorizontalAlignment = value; }
        }

        public DataTable DataSource
        {
            get { return dataSource; }
            set { dataSource = value; }
        }
        
        public XLBorderStyleValues BoderStyle
        {
            get { return boderStyle; }
            set { boderStyle = value; }
        }

        public string TitleSheet
        {
            get { return titleSheet; }
            set { titleSheet = value; }
        }

        public bool ShowGridLine
        {
            get { return _showGridLine; }
            set { _showGridLine = value; }
        }

        public string SheetName
        {
            get { return sheetName; }
            set { sheetName = value; }
        }

        public bool WrapText
        {
            get { return _wrapText; }
            set { _wrapText = value; }
        }
        public WorkSheetSetting Clone()
        {
            return (WorkSheetSetting)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }
}