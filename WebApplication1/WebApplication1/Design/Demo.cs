using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using WebApplication1.ExcelCore;

namespace WebApplication1.Design
{

    /// <summary>
    /// workbook manager
    /// </summary>
    public class BookDemo
    {
        private string _filename;
        List<WorksheetTemplateDemo> sheetTemplate;
        ClosedXML.Excel.XLWorkbook workbook;

        public BookDemo(string filename)
        {
            this.sheetTemplate = new List<WorksheetTemplateDemo>();
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

        public void AddSheet(WorksheetTemplateDemo sheet)
        {
            this.sheetTemplate.Add(sheet);
        }
    }

    public class WorksheetLayoutLevel1: WorksheetDemo
    {
        public WorksheetLayoutLevel1(WorksheetConfig sheetConfig)
            : base(sheetConfig)
        {
            this.SetStyle();
        }

        protected List<ExcelCore.ExcelColumnContent> columns;// = new List<ExcelColumnContent>();
        public void SetColumns(List<ExcelCore.ExcelColumnContent> _columns)
        {
            columns = _columns;
            this.setting.TableHeader.ColumnsWidth = _columns.ToArray();
        }
        protected ExcelStyle styleheader;
        protected ExcelStyle styleBase;
        protected ExcelStyle styletitle;
        protected ExcelStyle styleTableheader;
        protected ExcelStyle stylesubtitle;
        protected override void SetStyle()
        {
            WorksheetConfig setting = new WorksheetConfig();
            //set page
            setting.Content.BorderStyle = XLBorderStyleValues.Thin;
            setting.ShowGridLine = false;
            setting.Content.WrapText = false;
            setting.TableHeader.Style.Bold = false;
            setting.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            setting.PageSetup.PaperSize = XLPaperSize.A4Paper;

            //set style base
            styleBase = new ExcelStyle();
            styleBase.Horizontal = XLAlignmentHorizontalValues.Center;
            styleBase.Bold = true;

            //styleheader
            styleheader = styleBase.Clone();
            styleheader.Bold = false;

            //style title
            styletitle = styleBase.Clone();
            styletitle.Horizontal = XLAlignmentHorizontalValues.Center;
            styletitle.FontSize = 18;
            styletitle.FontColor = Color.Blue;
            
            //style table header
            styleTableheader = styleBase.Clone();
            styleTableheader.Horizontal = XLAlignmentHorizontalValues.Left;
            styleTableheader.Bold = true;
            setting.TableHeader.Style = styleTableheader;

            //style sub title
            stylesubtitle = styleBase.Clone();
            stylesubtitle.Horizontal = XLAlignmentHorizontalValues.Center;
            stylesubtitle.Bold = false;
            stylesubtitle.FontSize = 10;
            stylesubtitle.Italic = true;
            setting.Title.SubTitle.Style = stylesubtitle;
            
            setting.Title.Title.Style = styletitle;
            this.setting = setting;
        }

        public void SetName(string name)
        {
            this.setting.Name = name;
            this.name = name;
        }

        public void AddHeaderLeft(string text)
        {
            this.setting.Header.HeaderLeft.Add(new TextComponent { Text = text, Style = this.styleheader });
        }
        public void AddHeaderRight(string text)
        {
            this.setting.Header.HeaderRight.Add(new TextComponent { Text = text, Style = this.styleheader });
        }
        public void AddHeaderLeft(string text,ExcelStyle style)
        {
            this.setting.Header.HeaderLeft.Add(new TextComponent { Text = text, Style = style });
        }
        public void AddHeaderRight(string text, ExcelStyle style)
        {
            this.setting.Header.HeaderRight.Add(new TextComponent { Text = text, Style = style });
        }

        public void SetTitle(string title)
        {
            this.setting.Title.Title.Text = title;
        }
        public void SetSubTitle(string subTitle)
        {
            this.setting.Title.SubTitle.Text = subTitle;
        }

        public void SetDataSource(DataTable data)
        {
            this.setting.DataSource = data;
        }

        public WorksheetLayoutLevel1()
        {
            this.SetStyle();
        }
    }

    /// <summary>
    /// Sheet layout base
    /// </summary>
    public class WorksheetDemo : WorksheetTemplateDemo
    {
        public WorksheetDemo():base()
        {

        }
        public WorksheetDemo(WorksheetConfig sheetSetting)
            : base(sheetSetting)
        {
        }
        protected override void SetPage()
        {
            try
            {
                worksheet.PageSetup.PaperSize = this.setting.PageSetup.PaperSize;
                worksheet.PageSetup.PageOrientation = this.setting.PageSetup.PageOrientation;
                //worksheet.PageSetup.AdjustTo(100);
                worksheet.PageSetup.Margins.Left = this.setting.PageSetup.Margins.Left;
                worksheet.PageSetup.Margins.Right = this.setting.PageSetup.Margins.Right;
                worksheet.PageSetup.Margins.Top = this.setting.PageSetup.Margins.Top;
                worksheet.PageSetup.Margins.Bottom = this.setting.PageSetup.Margins.Bottom;
                worksheet.PageSetup.Margins.Header = this.setting.PageSetup.Margins.Header;
                worksheet.PageSetup.Margins.Footer = this.setting.PageSetup.Margins.Footer;
                worksheet.PageSetup.FitToPages(1, 1);
            }
            catch (Exception e)
            {
                Common.log.Error("Setting pagesetup", e);
            }
            
        }

        protected override void SetStyle()
        {
            try
            {
                worksheet.Style = this.worksheet.Style;
                worksheet.ShowGridLines = this.setting.ShowGridLine;
            }
            catch (Exception e)
            {
                Common.log.Error("Set Style", e);
            }
            
        }

        protected override void SetHeader()
        {
            try
            {
                int colLeft = Convert.ToInt32(setting.DataSource.Columns.Count * 2 / 5);
                int colSum = setting.DataSource.Columns.Count;
                int rowCount = this.setting.Header.HeaderLeft.Count > this.setting.Header.HeaderRight.Count ? this.setting.Header.HeaderLeft.Count : this.setting.Header.HeaderRight.Count;
                this.headerCells = new List<ExcelCell>();
                for (int i = 0; i < rowCount; i++)
                {
                    if (i < this.setting.Header.HeaderLeft.Count)
                    {
                        IXLCell cellDell = worksheet.Cell(this.row, 1);
                        worksheet.Range(this.row, 1, this.row, colLeft).Row(1).Merge();
                        cellDell.Value = this.setting.Header.HeaderLeft[i].Text;
                        this.setting.Header.HeaderLeft[i].Style.GetStyle(cellDell.Style);
                        this.headerCells.Add(new ExcelCell(cellDell, this.setting.Header.HeaderLeft[i].Style));
                    }
                    if (i < this.setting.Header.HeaderRight.Count)
                    {
                        IXLCell cellQuochieu = worksheet.Cell(this.row, colLeft + 1);
                        cellQuochieu.Value = this.setting.Header.HeaderRight[i].Text;
                        this.setting.Header.HeaderRight[i].Style.GetStyle(cellQuochieu.Style);
                        worksheet.Range(this.row, colLeft + 1, this.row, colSum).Row(1).Merge();
                        this.headerCells.Add(new ExcelCell(cellQuochieu, this.setting.Header.HeaderRight[i].Style));
                    }

                    this.row++;
                }
            }
            catch (Exception e)
            {

                Common.log.Error("Set header", e);
            }
        }

        protected override void SetSubHeader()
        {
            try
            {
                int colLeft = Convert.ToInt32(setting.DataSource.Columns.Count / 2);
                int colSum = setting.DataSource.Columns.Count;
                int rowCount = this.setting.Header.SubHeaderLeft.Count > this.setting.Header.SubHeaderRight.Count ? this.setting.Header.SubHeaderLeft.Count : this.setting.Header.SubHeaderRight.Count;
                if (this.headerCells == null || this.headerCells.Count == 0)
                {
                    this.headerCells = new List<ExcelCell>();
                }
                for (int i = 0; i < rowCount; i++)
                {
                    if (i < this.setting.Header.SubHeaderLeft.Count)
                    {
                        IXLCell cellDell = worksheet.Cell(this.row, 1);
                        worksheet.Range(this.row, 1, this.row, colLeft).Row(1).Merge();
                        cellDell.Value = this.setting.Header.HeaderLeft[i].Text;
                        this.setting.Header.HeaderLeft[i].Style.GetStyle(cellDell.Style);
                        this.headerCells.Add(new ExcelCell(cellDell, this.setting.Header.SubHeaderLeft[i].Style));
                    }
                    if (i < this.setting.Header.SubHeaderRight.Count)
                    {
                        IXLCell cellQuochieu = worksheet.Cell(this.row, colLeft + 1);
                        cellQuochieu.Value = this.setting.Header.HeaderRight[i].Text;
                        this.setting.Header.HeaderRight[i].Style.GetStyle(cellQuochieu.Style);
                        worksheet.Range(this.row, colLeft + 1, this.row, colSum).Row(1).Merge();
                        this.headerCells.Add(new ExcelCell(cellQuochieu, this.setting.Header.SubHeaderRight[i].Style));
                    }
                    this.row++;
                }
            }
            catch (Exception e)
            {
                Common.log.Error("Set subHeader", e);
            }
            
        }

        protected override void SetTitle()
        {
            try
            {
                int colSum = this.setting.DataSource.Columns.Count;
                int rowTitle = this.row + 1;
                var titleCell = worksheet.Cell(rowTitle, 1);
                this.titleCell = new ExcelCell(titleCell, this.setting.Title.Title.Style);
                titleCell.Value = this.setting.Title.Title.Text == null ? "" : this.setting.Title.Title.Text.ToUpper();
                this.setting.Title.Title.Style.GetStyle(titleCell.Style);
                worksheet.Range(rowTitle, 1, rowTitle, colSum).Row(1).Merge();
                this.row = rowTitle++;
            }
            catch (Exception e)
            {
                Common.log.Error("Set title", e);
            }
            
        }

        protected override void SetSubTitle()
        {
            try
            {
                int colSum = this.setting.DataSource.Columns.Count;
                int rowTitle = this.row + 1;
                var titleCell = worksheet.Cell(rowTitle, 1);
                this.subTitleCell = new ExcelCell(titleCell, this.setting.Title.SubTitle.Style);
                titleCell.Value = this.setting.Title.SubTitle.Text == null ? "" : this.setting.Title.SubTitle.Text;
                this.setting.Title.SubTitle.Style.GetStyle(titleCell.Style);
                worksheet.Range(rowTitle, 1, rowTitle, colSum).Row(1).Merge();
                this.row = rowTitle + 1;

            }
            catch (Exception e)
            {
                Common.log.Error("Set subtitle", e);
            }
        }

        protected override void SetContent()
        {
            try
            {
                
            if (this.setting.DataSource == null)
            {
                throw new NullReferenceException("Bạn chưa thiết lập dữ liệu");
            }
            else
            {
                int col = 1;
                var tag = worksheet.Cell(row, col).InsertData(this.setting.DataSource.AsEnumerable());
                this.tableRange.RangeAddress.LastAddress = tag.RangeAddress.LastAddress;
            }
            }
            catch (Exception e)
            {
                Common.log.Error("Set content", e);
            }
        }

        protected override void SetSignature()
        {

        }

        protected override void SetTableheader()
        {
            try
            {
                int colsum=this.setting.TableHeader.ColumnsWidth.Length;
                double width = this.setting.PageSetup.PageOrientation== XLPageOrientation.Landscape?128:85;
                this.setting.TableHeader.ColumnsWidth.Where(p => p.Width <= 0).ForEach(p => p.Width = width / colsum);
                if (this.setting.DataSource == null)
                {
                    throw new NullReferenceException("Bạn chưa thiết lập danh sách dữ liệu");
                }
                else
                {
                    if (this.setting.TableHeader.ColumnsWidth.Length != this.setting.DataSource.Columns.Count)
                    {
                        throw new ArgumentOutOfRangeException("Số lượng cột thiết lập không trùng với dataSource");
                    }
                    row++;
                    this.tableRange = worksheet.Range(this.row, 1, this.row, setting.DataSource.Columns.Count);
                    int col = 1;
                    foreach (DataColumn datacol in setting.DataSource.Columns)
                    {
                        IXLCell cell = worksheet.Cell(row, col);
                        this.setting.TableHeader.Style.GetStyle(cell.Style);
                        if (this.setting.TableHeader.ColumnsWidth == null)
                        {
                            cell.Value = datacol.ColumnName;
                        }
                        else
                        {
                            string name = this.setting.TableHeader.ColumnsWidth[col - 1].Caption;
                            worksheet.Cell(row, col).Value = (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(name)) ? name : datacol.ColumnName;
                            worksheet.Column(col).Width = this.setting.TableHeader.ColumnsWidth[col - 1].Width == 0 ? 8 : this.setting.TableHeader.ColumnsWidth[col - 1].Width;
                            worksheet.Column(col).Style.Alignment.Horizontal = this.setting.TableHeader.ColumnsWidth[col - 1].HorizontalAlignment;
                        }
                        col++;
                    }
                    if (this.setting.TableHeader.ColumnsWidth != null)
                    {
                        this.RefreshAlignmentHeader();
                    }
                    this.row++;
                }
            }
            catch (Exception e)
            {
                Common.log.Error("Set tableHeader",e);
            }
            
        }

        protected override void RefreshAlignmentHeader()
        {
            foreach (var item in this.headerCells)
            {
                item.Style.GetStyle(item.Cell.Style);
            }
            this.titleCell.Cell.Style.Alignment.Horizontal = this.titleCell.Style.Horizontal;
            this.subTitleCell.Cell.Style.Alignment.Horizontal = this.subTitleCell.Style.Horizontal;
        }

        protected override void SetBorderContent()
        {
            try
            {
                if (this.tableRange != null)
                {
                    tableRange.Cells().Style.Alignment.SetWrapText(this.setting.Content.WrapText);
                    this.tableRange.Style.Border.BottomBorder = this.setting.Content.BorderStyle;
                    this.tableRange.Style.Border.TopBorder = this.setting.Content.BorderStyle;
                    this.tableRange.Style.Border.LeftBorder = this.setting.Content.BorderStyle;
                    this.tableRange.Style.Border.RightBorder = this.setting.Content.BorderStyle;
                }
            }
            catch (Exception e)
            {
                Common.log.Error("Set border table", e);
            }
            
        }
    }

    /// <summary>
    /// Worksheet template
    /// </summary>
    public abstract class WorksheetTemplateDemo
    {
        protected string name;
        protected WorksheetConfig setting;
        protected IXLWorksheet worksheet;
        protected int row = 1;
        protected int totalColumns = 0;
        protected List<ExcelCell> headerCells = new List<ExcelCell>();
        protected IXLRange tableRange;
        protected ExcelCell titleCell;
        protected ExcelCell subTitleCell;
        protected IXLCell cellSubtitle;

        public WorksheetTemplateDemo(){}
        public WorksheetTemplateDemo(WorksheetConfig sheetSetting)
        {
            this.setting = sheetSetting;
            this.name = sheetSetting.Name;
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
        public WorksheetConfig Setting
        {
            get { return setting; }
            set { setting = value; }
        }
        public IXLWorksheet GetTemplate()
        {
            try
            {
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
            }
            catch (Exception e)
            {
                Common.log.Error("Get Template",e);
            }
            
            return this.worksheet;
        }
        public IXLWorksheet GetTemplate(IXLWorksheet sheet)
        {
            try
            {
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
            }
            catch (Exception e)
            {
                Common.log.Error("Get template from worksheet", e);
            }
            return this.worksheet;
        }
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