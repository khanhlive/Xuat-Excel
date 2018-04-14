using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication1.ExcelCore
{
    public class ExcelSimple
    {
        private XLPaperSize paperSize= XLPaperSize.A4Paper;

        private XLPageOrientation pageOrientation = XLPageOrientation.Default;

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

        private bool tableHeaderBold;

        public bool TableHeaderBold
        {
            get { return tableHeaderBold; }
            set { tableHeaderBold = value; }
        }

        private ExcelColumnContent[] columnsWidth;

        public ExcelColumnContent[] ColumnsWidth
        {
            get { return columnsWidth; }
            set { columnsWidth = value; }
        }

        private XLAlignmentHorizontalValues headerHorizontalAlignment;

        public XLAlignmentHorizontalValues HeaderHorizontalAlignment
        {
            get { return headerHorizontalAlignment; }
            set { headerHorizontalAlignment = value; }
        }

        private XLBorderStyleValues boderStyle;
        protected bool _wrapText = true;
        protected string sheetsName;
        protected bool _showGridLine = false;
        protected IXLWorksheet worksheet;
        protected DataTable dataSource;
        private List<IXLCell> headerCells = new List<IXLCell>();

        public DataTable DataSource
        {
            get { return dataSource; }
            set { dataSource = value; }
        }
        private string titleSheet;
        private int row ;
        protected IXLRange tableRange;

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

        public ExcelSimple(IXLWorksheet worksheet)
        {
            this.worksheet = worksheet;
            this.row = 1;
        }

        public bool ShowGridLine
        {
            get { return _showGridLine; }
            set { _showGridLine = value; }
        }

        public string SheetsName
        {
            get { return sheetsName; }
            set { sheetsName = value; }
        }

        public bool WrapText
        {
            get { return _wrapText; }
            set { _wrapText = value; }
        }

        protected virtual void SetWorkSheet()
        {
            
            worksheet.Style.Font.FontName = "Times New Roman";
            worksheet.Style.Font.FontSize = 13;
            worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.ShowGridLines = this.ShowGridLine;
            
        }

        protected virtual void SetHeader()
        {
            //string tenpb = ddlPhongBan.SelectedItem == null ? "Tiêu đề demo" : ddlPhongBan.SelectedItem.Text;
            string tenchuquan = "Sở GD-ĐT Hà Nội";
            string tendonvi = "Trường THPT Ba Đình";
            string quochieu = "Cộng hòa xã hội chủ nghĩa việt nam";
            string tieungu = "Độc lập - Tự do - Hạnh phúc";
            int colLeft = Convert.ToInt32(dataSource.Columns.Count * 2 / 5);
            int colSum = dataSource.Columns.Count;

            IXLCell cellChuquan = worksheet.Cell(1, 1);
            worksheet.Range(this.row, 1, this.row, colLeft).Row(1).Merge();
            IXLCell cellQuochieu = worksheet.Cell(1, colLeft + 1);
            worksheet.Range(this.row, colLeft + 1, this.row, colSum).Row(1).Merge();
            this.row++;
            IXLCell cellDonvi = worksheet.Cell(this.row, 1);
            worksheet.Range(this.row, 1, this.row, colLeft).Row(1).Merge();
            
            IXLCell cellTieungu = worksheet.Cell(this.row, colLeft + 1);
            worksheet.Range(this.row, colLeft + 1, this.row, colSum).Row(1).Merge();
            this.row++;
            cellChuquan.Value = tenchuquan.ToUpper();
            //cellChuquan.Style.Font.Bold = true;
            cellChuquan.Style.Alignment.Horizontal = this.headerHorizontalAlignment; ;

            cellDonvi.Value = tendonvi.ToUpper();
            cellDonvi.Style.Alignment.Horizontal =  this.headerHorizontalAlignment; ;
            cellDonvi.Style.Font.Bold = true;

            cellQuochieu.Value = quochieu.ToUpper();
            cellQuochieu.Style.Alignment.Horizontal =  this.headerHorizontalAlignment; ;
            cellQuochieu.Style.Font.Bold = true;

            cellTieungu.Value = tieungu;
            cellTieungu.Style.Alignment.Horizontal = this.headerHorizontalAlignment; ;
            cellTieungu.Style.Font.Bold = true;
            this.headerCells = new List<IXLCell>();
            this.headerCells.Add(cellChuquan);
            this.headerCells.Add(cellDonvi);
            this.headerCells.Add(cellQuochieu);
            this.headerCells.Add(cellTieungu);
        }

        protected virtual void SetTableHeader()
        {
            if (this.dataSource == null)
            {
                throw new NullReferenceException("Bạn chưa thiết lập danh dữ liệu");
            }
            else
            {
                if (this.ColumnsWidth.Length != this.dataSource.Columns.Count)
                {
                    throw new ArgumentOutOfRangeException("Số lượng cột thiết lập không trùng với dataSource");
                }
                this.tableRange = worksheet.Range(this.row, 1, this.row, dataSource.Columns.Count);
                int col = 1;
                foreach (DataColumn datacol in dataSource.Columns)
                {
                    worksheet.Cell(row, col).Style.Font.Bold = this.tableHeaderBold;
                    if (this.columnsWidth==null)
                    {
                        worksheet.Cell(row, col).Value = datacol.ColumnName;
                    }
                    else
                    {
                        string name = this.columnsWidth[col - 1].Name;
                        worksheet.Cell(row, col).Value = (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(name)) ? name: datacol.ColumnName;
                        worksheet.Column(col).Width = this.columnsWidth[col - 1].Width == 0 ? 8 : this.columnsWidth[col - 1].Width;
                        worksheet.Column(col).Style.Alignment.Horizontal = this.columnsWidth[col - 1].HorizontalAlignment;
                    }
                    col++;
                }
                if (this.columnsWidth!=null)
                {
                    this.RefreshAlignmentHeader();
                }
                this.row++;
            }
        }

        protected virtual void SetData()
        {
            if (this.dataSource == null)
            {
                throw new NullReferenceException("Bạn chưa thiết lập danh dữ liệu");
            }
            else
            {
                int col = 1;
                var tag = worksheet.Cell(row, col).InsertData(this.dataSource.AsEnumerable());
                this.tableRange.RangeAddress.LastAddress = tag.RangeAddress.LastAddress;
            }
        }

        private IXLCell titleCell;

        protected virtual void SetTitle()
        {
            int colSum = this.dataSource.Columns.Count;
            int rowTitle = this.row+1;
            var titleCell = worksheet.Cell(rowTitle, 1);
            this.titleCell = titleCell;
            titleCell.Value = this.titleSheet == null ? "" : this.titleSheet.ToUpper();
            titleCell.Style.Font.FontSize = 18;
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontColor = XLColor.Blue;
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range(rowTitle, 1, rowTitle, colSum).Row(1).Merge();
            this.row = rowTitle+2;
        }

        protected virtual void SetWorkSheetPageSetup()
        {
            
            worksheet.PageSetup.PaperSize = this.paperSize;
            worksheet.PageSetup.PageOrientation = this.pageOrientation;
            worksheet.PageSetup.AdjustTo(100);
            worksheet.PageSetup.Margins.Left = 0.7;
            worksheet.PageSetup.Margins.Right = 0.7;
            worksheet.PageSetup.Margins.Top = 0.5;
            worksheet.PageSetup.Margins.Bottom = 0;
            worksheet.PageSetup.Margins.Header = 0.5;
            worksheet.PageSetup.Margins.Footer = 0.5;
        }

        protected virtual void RefreshAlignmentHeader()
        {
            foreach (var item in this.headerCells)
            {
                item.Style.Alignment.Horizontal = this.headerHorizontalAlignment;
            }
            this.titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        public virtual IXLWorksheet Init()
        {
            this.SetWorkSheet();
            this.SetWorkSheetPageSetup();
            this.SetHeader();
            this.SetTitle();
            this.SetTableHeader();
            this.SetData();
            this.SetBorderTable();
            return this.worksheet;
            
        }

        private void SetBorderTable()
        {
            if (this.tableRange != null)
            {
                tableRange.Cells().Style.Alignment.SetWrapText(this._wrapText);
                this.tableRange.Style.Border.TopBorder = this.boderStyle;
                this.tableRange.Style.Border.BottomBorder = this.boderStyle;
                this.tableRange.Style.Border.LeftBorder = this.boderStyle;
                this.tableRange.Style.Border.RightBorder = this.boderStyle;
            }
        }

        public void Dispose()
        {
            if (this.worksheet != null)
                this.worksheet.Dispose();
        }
    }
}