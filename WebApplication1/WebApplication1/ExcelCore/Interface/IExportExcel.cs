using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApplication1.ExcelCore.Interface
{
    public interface IExportExcel
    {
        void Export();
        void AddSheet(WorkSheetSetting settings);
    }

    public class WorksheetManager
    {
        protected string name;
        protected WorkSheetSetting setting;
        protected IXLWorksheet worksheet;
        protected int row = 1;
        protected List<IXLCell> headerCells = new List<IXLCell>();
        protected IXLRange tableRange;
        protected IXLCell titleCell;
        public WorksheetManager(IXLWorksheet sheet)
        {
            this.worksheet = sheet;
        }

        public WorkSheetSetting Setting
        {
            get { return setting; }
            set { setting = value; }
        }

        public ClosedXML.Excel.IXLWorksheet RenderSheet()
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

        protected virtual void SetWorkSheet()
        {
            worksheet.Style.Font.FontName = "Times New Roman";
            worksheet.Style.Font.FontSize = 13;
            worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.ShowGridLines = this.setting.ShowGridLine;

        }

        protected virtual void SetHeader()
        {
            string tenchuquan = "Sở GD-ĐT Hà Nội";
            string tendonvi = "Trường THPT Ba Đình";
            string quochieu = "Cộng hòa xã hội chủ nghĩa việt nam";
            string tieungu = "Độc lập - Tự do - Hạnh phúc";
            int colLeft = Convert.ToInt32(setting.DataSource.Columns.Count * 2 / 5);
            int colSum = setting.DataSource.Columns.Count;

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
            cellChuquan.Style.Alignment.Horizontal = this.setting.HeaderHorizontalAlignment; ;

            cellDonvi.Value = tendonvi.ToUpper();
            cellDonvi.Style.Alignment.Horizontal = this.setting.HeaderHorizontalAlignment; ;
            cellDonvi.Style.Font.Bold = true;

            cellQuochieu.Value = quochieu.ToUpper();
            cellQuochieu.Style.Alignment.Horizontal = this.setting.HeaderHorizontalAlignment; ;
            cellQuochieu.Style.Font.Bold = true;

            cellTieungu.Value = tieungu;
            cellTieungu.Style.Alignment.Horizontal = this.setting.HeaderHorizontalAlignment; ;
            cellTieungu.Style.Font.Bold = true;
            this.headerCells = new List<IXLCell>();
            this.headerCells.Add(cellChuquan);
            this.headerCells.Add(cellDonvi);
            this.headerCells.Add(cellQuochieu);
            this.headerCells.Add(cellTieungu);
        }

        protected virtual void SetTableHeader()
        {
            if (this.setting.DataSource == null)
            {
                throw new NullReferenceException("Bạn chưa thiết lập danh dữ liệu");
            }
            else
            {
                if (this.setting.ColumnsWidth.Length != this.setting.DataSource.Columns.Count)
                {
                    throw new ArgumentOutOfRangeException("Số lượng cột thiết lập không trùng với dataSource");
                }
                this.tableRange = worksheet.Range(this.row, 1, this.row, setting.DataSource.Columns.Count);
                int col = 1;
                foreach (DataColumn datacol in setting.DataSource.Columns)
                {
                    worksheet.Cell(row, col).Style.Font.Bold = this.setting.TableHeaderBold;
                    if (this.setting.ColumnsWidth == null)
                    {
                        worksheet.Cell(row, col).Value = datacol.ColumnName;
                    }
                    else
                    {
                        string name = this.setting.ColumnsWidth[col - 1].Name;
                        worksheet.Cell(row, col).Value = (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(name)) ? name : datacol.ColumnName;
                        worksheet.Column(col).Width = this.setting.ColumnsWidth[col - 1].Width == 0 ? 8 : this.setting.ColumnsWidth[col - 1].Width;
                        worksheet.Column(col).Style.Alignment.Horizontal = this.setting.ColumnsWidth[col - 1].HorizontalAlignment;
                    }
                    col++;
                }
                if (this.setting.ColumnsWidth != null)
                {
                    this.RefreshAlignmentHeader();
                }
                this.row++;
            }
        }

        protected virtual void SetData()
        {
            if (this.setting.DataSource == null)
            {
                throw new NullReferenceException("Bạn chưa thiết lập danh dữ liệu");
            }
            else
            {
                int col = 1;
                var tag = worksheet.Cell(row, col).InsertData(this.setting.DataSource.AsEnumerable());
                this.tableRange.RangeAddress.LastAddress = tag.RangeAddress.LastAddress;
            }
        }

        protected virtual void SetTitle()
        {
            int colSum = this.setting.DataSource.Columns.Count;
            int rowTitle = this.row + 1;
            var titleCell = worksheet.Cell(rowTitle, 1);
            this.titleCell = titleCell;
            titleCell.Value = this.setting.TitleSheet == null ? "" : this.setting.TitleSheet.ToUpper();
            titleCell.Style.Font.FontSize = 18;
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontColor = XLColor.Blue;
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range(rowTitle, 1, rowTitle, colSum).Row(1).Merge();
            this.row = rowTitle + 2;
        }

        protected virtual void SetWorkSheetPageSetup()
        {

            worksheet.PageSetup.PaperSize = this.setting.PaperSize;
            worksheet.PageSetup.PageOrientation = this.setting.PageOrientation;
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
                item.Style.Alignment.Horizontal = this.setting.HeaderHorizontalAlignment;
            }
            this.titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        protected virtual void SetBorderTable()
        {
            if (this.tableRange != null)
            {
                tableRange.Cells().Style.Alignment.SetWrapText(this.setting.WrapText);
                this.tableRange.Style.Border.TopBorder = this.setting.BoderStyle;
                this.tableRange.Style.Border.BottomBorder = this.setting.BoderStyle;
                this.tableRange.Style.Border.LeftBorder = this.setting.BoderStyle;
                this.tableRange.Style.Border.RightBorder = this.setting.BoderStyle;
            }
        }

        public virtual string GetName()
        {
            return this.name;
        }
    }
}
