using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApplication1.ExcelCore;

namespace WebApplication1.Design
{
    public abstract class WorksheetTemplate
    {
        protected string name;
        protected WorkSheetSetting setting;
        protected IXLWorksheet worksheet;
        protected int row = 1;
        protected int totalColumns=0;
        protected List<IXLCell> headerCells = new List<IXLCell>();
        protected IXLRange tableRange;
        protected IXLCell titleCell;
        protected IXLCell cellSubtitle;

        public WorksheetTemplate( WorkSheetSetting sheetSetting)
        {
            this.setting = sheetSetting;
            this.name = sheetSetting.SheetName;
            if (this.setting.DataSource != null)
            {
                this.totalColumns = this.setting.DataSource.Columns.Count;
            }
        }

        public void SetWorksheet(IXLWorksheet sheet)
        {
            this.worksheet = sheet;
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public WorkSheetSetting Setting
        {
            get { return setting; }
            set { setting = value; }
        }
        public IXLWorksheet GetTemplate()
        {
            //this.CreateSheet();
            this.SetPage();
            this.SetStyle();
            this.SetHeader();
            this.SetSubHeader();
            this.SetTitle();
            this.SetSubTitle();
            this.SetTableheader();
            this.SetContent();
            this.RefreshAlignmentHeader();
            this.SetBorderContent();
            this.SetSignature();
            return this.worksheet;
        }
        public IXLWorksheet GetTemplate(IXLWorksheet sheet)
        {
            //this.CreateSheet();
            this.worksheet = sheet;
            this.SetPage();
            this.SetStyle();
            this.SetHeader();
            this.SetSubHeader();
            this.SetTitle();
            this.SetSubTitle();
            this.SetTableheader();
            this.SetContent();
            this.RefreshAlignmentHeader();
            this.SetBorderContent();
            this.SetSignature();
            return this.worksheet;
        }
        //protected abstract void CreateSheet();
        protected abstract void SetPage();
        protected abstract void SetStyle();
        protected abstract void SetHeader();
        protected abstract void SetSubHeader();
        protected abstract void SetTitle();
        protected abstract void SetSubTitle();
        protected abstract void SetTableheader();
        protected abstract void SetContent();
        protected abstract void SetBorderContent();
        protected abstract void RefreshAlignmentHeader();
        protected abstract void SetSignature();
        

    }
}
