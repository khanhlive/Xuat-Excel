
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication1.Providers
{
    /// <summary>
    /// Báo cáo Hồ Chí Minh Môn nhận xét
    /// </summary>
    public class BaoCaoHCM_C2RenderMonNhanXet : BaoCaoHCMExportDiemMonhoc
    {
        public BaoCaoHCM_C2RenderMonNhanXet(string tenmonhoc, string namhoc) : base(tenmonhoc,namhoc) { }

        public override void Export(string fileTemplate,string name,DataTable[] dataSource, int hangBatDauBangChiTiet, int hangBatDauBangThongKe )
        {
            try
            {
                using (XLWorkbook workbook = new XLWorkbook(fileTemplate))
                {
                    if (dataSource.Length==3)
                    {
                        #region lấy danh sách worksheet và thiết lập tiêu đề môn học
                        var sheet = workbook.Worksheet(1);
                        sheet.Cell(1, 5).Value = this.TieudeThiHK;
                        var sheet2 = workbook.Worksheet(2);
                        sheet.Cell(1, 5).Value = this.TieudeXLHK2;
                        var sheet3 = workbook.Worksheet(3);
                        sheet.Cell(1, 5).Value = this.TieudeXLCN;
                        #endregion

                        #region Thiết lập danh sách cột dữ liệu
                        List<ColumnDetail> columns = new List<ColumnDetail>();
                        columns.Add(new ColumnDetail { FieldName = "TEN_TRUONG" });
                        columns.Add(new ColumnDetail { FieldName = "SO_LUONGHS" });
                        columns.Add(new ColumnDetail { FieldName = "DAT_TS" });
                        columns.Add(new ColumnDetail { FieldName = "DAT_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "CHUADAT_TS" });
                        columns.Add(new ColumnDetail { FieldName = "CHUADAT_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "TBTROLEN_TS" });
                        columns.Add(new ColumnDetail { FieldName = "TBTROLEN_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "TBTROLEN_XEP_HANG" });
                        #endregion
                        
                        ////thiết lập xuất dữ liệu theo các sheet tương ứng

                        this.GetWorkSheet(dataSource[0], sheet, columns, hangBatDauBangChiTiet, hangBatDauBangThongKe);
                        this.GetWorkSheet(dataSource[1], sheet2, columns, hangBatDauBangChiTiet, hangBatDauBangThongKe);
                        this.GetWorkSheet(dataSource[2], sheet3, columns, hangBatDauBangChiTiet, hangBatDauBangThongKe);
                    }
                    #region Response
                    this.Response(workbook, name);
                    #endregion

                }
            }
            catch (Exception e)
            {
                Common.log.Error("error", e);
            }
            finally
            {
                if (System.IO.File.Exists(fileTemplate))
                    System.IO.File.Delete(fileTemplate);
            }
            
        }

        
        protected void GetWorkSheet(DataTable dataSource, IXLWorksheet sheet, List<ColumnDetail> columns, int rowStart, int rowToanTruongStart)
        {
            DataView dview = new DataView(dataSource, "KHOI = '6' OR KHOI ='60'", "KHOI,TEN_TRUONG", DataViewRowState.CurrentRows);
            List<int> rowIndexGroup = new List<int>();
            DataTable khoi6 = dview.ToTable(false, columns.Select(p => p.FieldName).ToArray());
            DataTable khoi7 = new DataView(dataSource, "KHOI = '7' OR KHOI ='70'", "KHOI,TEN_TRUONG", DataViewRowState.CurrentRows).ToTable(false, columns.Select(p => p.FieldName).ToArray());
            DataTable khoi8 = new DataView(dataSource, "KHOI = '8' OR KHOI ='80'", "KHOI,TEN_TRUONG", DataViewRowState.CurrentRows).ToTable(false, columns.Select(p => p.FieldName).ToArray());
            DataTable khoi9 = new DataView(dataSource, "KHOI = '9' OR KHOI ='90'", "KHOI,TEN_TRUONG", DataViewRowState.CurrentRows).ToTable(false, columns.Select(p => p.FieldName).ToArray());
            int row = rowStart; 
            sheet.Column(1).Width = 20;
            int rowBegin = Convert.ToInt32(row.ToString());
            row = this.SetGroup(sheet, khoi6, row, columns);
            rowIndexGroup.Add(row - 1);
            row = this.SetGroup(sheet, khoi7, row, columns);
            rowIndexGroup.Add(row - 1);
            row = this.SetGroup(sheet, khoi8, row, columns);
            rowIndexGroup.Add(row - 1);
            row = this.SetGroup(sheet, khoi9, row, columns);
            rowIndexGroup.Add(row - 1);
            SetGroupTHCS(sheet, rowIndexGroup, row, columns);


            //set border table
            var range = sheet.Range(rowBegin, 1, row, columns.Count);
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            int rowEnd = row + rowToanTruongStart - rowBegin;
            List<int> chisoBangToanTruong = SetToanTruong(sheet, rowIndexGroup, ref rowEnd, columns);
            SetGroupTHCS(sheet, chisoBangToanTruong, rowEnd, columns);
            var rangeToanTruong = sheet.Range(row + rowToanTruongStart - rowBegin, 1, rowEnd, columns.Count);
            rangeToanTruong.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            rangeToanTruong.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            rangeToanTruong.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            rangeToanTruong.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        }


        /// <summary>
        /// In bảng thống kê toàn trường
        /// </summary>
        /// <param name="sheet">sheet</param>
        /// <param name="indexs">danh sách index hàng chứa tổng hợp các khối</param>
        /// <param name="rowStart">chỉ số hàng bắt đầu in</param>
        /// <param name="columns">danh sách cột</param>
        /// <returns>chỉ số hàng tiếp theo có thể được In tiếp dữ liệu</returns>
        protected List<int> SetToanTruong(IXLWorksheet sheet, List<int> indexs,ref int rowStart, List<ColumnDetail> columns)
        {
            List<int> range = new List<int>();
            sheet.Row(rowStart).InsertRowsBelow(indexs.Count+1);
            int rowBegin = Convert.ToInt32(rowStart.ToString());
            string[] khoi = new string[] { "6", "7", "8", "9" };
            for (int j = 0; j < indexs.Count; j++)
            {
                for (int i = 1; i <= columns.Count; i++)
                {
                    var cell = sheet.Cell(rowStart, i);
                    cell.Style.Font.Bold = false;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    #region Thiết lập cột
                    switch (columns[i - 1].FieldName)
                    {
                        case "DAT_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 3).Address.ToString());
                            break;
                        case "CHUADAT_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 5).Address.ToString());
                            break;
                        case "SO_LUONGHS":
                            cell.FormulaA1 = string.Format("=SUM({0},{1})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 5).Address.ToString());
                            break;
                        case "DAT_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            cell.Style.NumberFormat.NumberFormatId = 10;
                            break;
                        case "CHUADAT_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            cell.Style.NumberFormat.NumberFormatId = 10;
                            break;
                        case "TBTROLEN_TS":
                            cell.FormulaA1 = string.Format("=SUM({0})", sheet.Cell(rowStart, 3).Address.ToString());
                            break;
                        case "TBTROLEN_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 7).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            cell.Style.NumberFormat.NumberFormatId = 10;
                            break;
                        case "TBTROLEN_XEP_HANG":
                            cell.FormulaA1 = string.Format("=RANK({0},{1}:{2},0)", sheet.Cell(rowStart, 8).Address.ToString(), this.GetAddressString(sheet.Cell(rowBegin, 8), true), this.GetAddressString(sheet.Cell(rowBegin + khoi.Length - 1, 8), true));
                            break;
                        case "TEN_TRUONG":
                            cell.Value = khoi[j];
                            break;
                        default:
                            break;
                    }
                    #endregion


                    //set format string cho cột %
                    if (columns[i-1].DataType == XLCellValues.Number)
                    {
                        cell.DataType = columns[i - 1].DataType;
                        cell.Style.NumberFormat.NumberFormatId = 10;
                    }
                }
                range.Add(rowStart);
                rowStart++;
            }
            return range;
        }

        /// <summary>
        /// In thống kê toàn trường
        /// </summary>
        /// <param name="sheet">worksheet</param>
        /// <param name="indexs">danh sách index hàng chứa tổng hợp các khối</param>
        /// <param name="rowStart">chỉ số hàng bắt đầu in</param>
        /// <param name="columns">danh sách cột</param>
        protected void SetGroupTHCS(IXLWorksheet sheet, List<int> indexs, int rowStart, List<ColumnDetail> columns)
        {
            List<string> cellsDatTS = new List<string>();
            List<string> cellsChuaDatTS = new List<string>();
            foreach (var item in indexs)
            {
                cellsDatTS.Add( sheet.Cell(item, 3).Address.ToString());
                cellsChuaDatTS.Add(sheet.Cell(item, 5).Address.ToString());
            }
            for (int i = 1; i <= columns.Count; i++)
            {
                var cell = sheet.Cell(rowStart, i);
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                #region Thiết lập cột
                switch (columns[i - 1].FieldName)
                {
                    case "DAT_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsDatTS));
                        break;
                    case "CHUADAT_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsChuaDatTS));
                        break;
                    case "SO_LUONGHS":
                        cell.FormulaA1 = string.Format("=SUM({0},{1})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 5).Address.ToString());
                        break;
                    case "DAT_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        cell.Style.NumberFormat.NumberFormatId = 10;
                        break;
                    case "CHUADAT_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        cell.Style.NumberFormat.NumberFormatId = 10;
                        break;
                    case "TBTROLEN_TS":
                        cell.FormulaA1 = string.Format("=SUM({0})", sheet.Cell(rowStart, 3).Address.ToString());
                        break;
                    case "TBTROLEN_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 7).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        cell.Style.NumberFormat.NumberFormatId = 10;
                        break;
                    case "TBTROLEN_XEP_HANG":
                        break;
                    case "TEN_TRUONG":
                        cell.Value = "THCS";
                        break;
                    default:
                        break;
                }
                #endregion
                
            }
        }

        /// <summary>
        /// In thống kê các trường nhóm theo khối
        /// </summary>
        /// <param name="sheet">worksheet</param>
        /// <param name="dt">Dữ liệu nguồn</param>
        /// <param name="rowStart">chỉ số hành bắt đầu IN</param>
        /// <param name="columns">danh sách các cột</param>
        /// <returns>chỉ số hàng tiếp theo có thể được In tiếp dữ liệu</returns>
        protected int SetGroup(IXLWorksheet sheet, DataTable dt, int rowStart, List<ColumnDetail> columns)
        {
            int colSum = columns.Count;
            var rowInsert = sheet.Range(rowStart, 1, rowStart, colSum);
            rowInsert.InsertRowsBelow(dt.Rows.Count);
            int row = Convert.ToInt32(rowStart.ToString());
            for (int ro = 0; ro < dt.Rows.Count; ro++)
            {
                DataRow item = dt.Rows[ro];
                int i = 1;
                foreach (var col in columns)
                {
                    var cell = sheet.Cell(row, i);
                    if (ro == dt.Rows.Count - 1)
                        cell.Style.Font.Bold = true;
                    else cell.Style.Font.Bold = false;
                    if (i==1)
                    {
                        //tên trường
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }
                    else //cột thường
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    #region Thiết lập cột
                    switch (col.FieldName)
                    {
                        case "DAT_TS": if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(row - 1, 3).Address.ToString());
                            else
                                cell.Value = item[col.FieldName];
                            break;
                        case "CHUADAT_TS":
                            if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(row - 1, 5).Address.ToString());
                            else
                                cell.Value = item[col.FieldName];
                            break;
                        case "SO_LUONGHS":
                            cell.FormulaA1 = string.Format("=SUM({0},{1})", sheet.Cell(row, 3).Address.ToString(), sheet.Cell(row, 5).Address.ToString());
                            break;
                        case "TEN_TRUONG":
                            cell.Value = item[col.FieldName];
                            if (ro == dt.Rows.Count - 1)
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            break;
                        case "DAT_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 3).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "CHUADAT_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 5).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "TBTROLEN_TS":
                            cell.FormulaA1 = string.Format("=SUM({0})", sheet.Cell(row, 3).Address.ToString());
                            break;
                        case "TBTROLEN_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 7).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "TBTROLEN_XEP_HANG":
                            if (ro != dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=RANK({0},{1}:{2},0)", sheet.Cell(row, 8).Address.ToString(), this.GetAddressString(sheet.Cell(rowStart, 8),true), this.GetAddressString(sheet.Cell(rowStart + dt.Rows.Count - 2, 8),true));
                            break;
                        default: cell.Value = item[col.FieldName];
                            break;
                    }
                    #endregion
                    
                    //set cell type
                    if (col.DataType == XLCellValues.Number)
                    {
                        cell.DataType = col.DataType;
                        cell.Style.NumberFormat.NumberFormatId = 10;
                    }
                    i++;
                }
                row++;
            }
            return row;
        }

    }

    

    public class ColumnDetail
    {
        public string FieldName { get; set; }
        public string FormulaA1 { get; set; }
        public XLCellValues DataType { get; set; }
    }
}