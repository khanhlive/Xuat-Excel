using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication1.Providers
{
    public class BaoCaoHCMManager
    {
        /// <summary>
        /// Báo cáo HCM: Xuất báo cáo thống kê điểm số các môn: bao gồm các môn nhận xét và điểm số 
        /// </summary>
        /// <param name="fileTemplate">Đường dẫn file template đã Upload</param>
        /// <param name="dataSource">Danh sách dữ liệu nguồn: bao gồm 3 sheet, mỗi sheet một bảng dữ liệu</param>
        /// <param name="hangBatDauBangChiTiet">Chỉ số hàng bắt đầu In thống kê chi tiết(theo hàng trong file template)</param>
        /// <param name="hangBatDauBangToantruong">Chỉ số hàng bắt đầu In thống kê toàn trường(theo hàng trong file template)</param>
        public static void XuatBaoCaoHCM(string fileTemplate, string tenmonhoc,string namhoc, DataTable[] dataSource, int hangBatDauBangChiTiet, int hangBatDauBangToantruong)
        {
            try
            {
                if (dataSource.Length > 0)
                {
                    BaoCaoHCMExportDiemMonhoc baocaoHCM;
                    if (dataSource[0].Columns.Count>10)
                    {
                        // In báo cáo môn điểm số
                        baocaoHCM = new BaoCaoHCM_C2RenderMonDiemSo(tenmonhoc,namhoc);
                        
                    }
                    else
                    {
                        // In báo cáo môn nhận xét
                        baocaoHCM = new BaoCaoHCM_C2RenderMonNhanXet(tenmonhoc, namhoc);
                    }
                    baocaoHCM.Export(fileTemplate, tenmonhoc, dataSource, hangBatDauBangChiTiet, hangBatDauBangToantruong);
                }
            }
            catch (Exception)
            {
                
                throw;
            }
        }


        /// <summary>
        /// Báo cáo HCM: Xuất báo cáo thống kê điểm số các môn: bao gồm các môn nhận xét và điểm số 
        /// </summary>
        /// <param name="fileTemplate">Đường dẫn file template đã Upload</param>
        /// <param name="MaPGD">Tham số: Mã phòng</param>
        /// <param name="MaNamhoc">Tham số: Mã năm học</param>
        /// <param name="MaMonhoc">Tham số: Mã môn học</param>
        /// <param name="Hocky">Tham số: Học kỳ</param>
        /// <param name="SoLieu">Tham số: Số liệu</param>
        public static void XuatBaoCaoHCM(string fileTemplate, string tenmonhoc,string namhoc, string MaPGD, string MaNamhoc, string MaMonhoc, int Hocky, int SoLieu)
        {
            try
            {
                BaoCaoHCM_C2 baocaohcm = new BaoCaoHCM_C2();
                DataTable tableThiHK = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamhoc, MaMonhoc, Hocky, SoLieu);
                SoLieu = 2;
                DataTable tableXLHK = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamhoc, MaMonhoc, Hocky, SoLieu);
                Hocky = 3;
                DataTable tableXLCN = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamhoc, MaMonhoc, Hocky, SoLieu);
                BaoCaoHCMExportDiemMonhoc baocaoHCM;
                if (tableThiHK.Columns.Count > 10)
                {
                    // In báo cáo môn điểm số
                    baocaoHCM = new BaoCaoHCM_C2RenderMonDiemSo(tenmonhoc,namhoc);
                }
                else
                {
                    // In báo cáo môn nhận xét
                    baocaoHCM = new BaoCaoHCM_C2RenderMonNhanXet(tenmonhoc, namhoc);
                }
                baocaoHCM.Export(fileTemplate, tenmonhoc, new DataTable[] { tableThiHK, tableXLHK, tableXLCN }, 6, 12);

            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void XuatBaocaoHCM_Nhieumon(string fileTemplateMonNhanXet, string fileTemplateMonDiemSo, List<MonHoc> monhoc,string MaPGD, string MaNamhoc, int Hocky, int SoLieu)
        {
            foreach (var item in monhoc)
            {
                try
                {
                    BaoCaoHCM_C2 baocaohcm = new BaoCaoHCM_C2();
                    DataTable tableThiHK = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamhoc, item.MaMonHoc, Hocky, SoLieu);
                    SoLieu = 2;
                    DataTable tableXLHK = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamhoc, item.MaMonHoc, Hocky, SoLieu);
                    Hocky = 3;
                    DataTable tableXLCN = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamhoc, item.MaMonHoc, Hocky, SoLieu);
                    BaoCaoHCMExportDiemMonhoc baocaoHCM;
                    string filePath = "";
                    if (tableThiHK.Columns.Count > 10)
                    {
                        // In báo cáo môn điểm số
                        baocaoHCM = new BaoCaoHCM_C2RenderMonDiemSo(item.TenMonHoc,"2017 - 2018");
                        filePath = fileTemplateMonDiemSo;
                    }
                    else
                    {
                        // In báo cáo môn nhận xét
                        baocaoHCM = new BaoCaoHCM_C2RenderMonNhanXet(item.TenMonHoc, "2017 - 2018");
                        filePath = fileTemplateMonNhanXet;
                    }
                    baocaoHCM=new BaoCaoHCM_C2RenderMonDiemSo(item.TenMonHoc,"2017 - 2018");
                    baocaoHCM.Export(filePath, item.TenMonHoc,new DataTable[] { tableThiHK, tableXLHK, tableXLCN }, 6, 12);
                    //baocaoHCM.Export(filePath, item.TenMonHoc, new DataTable[] { dt,dt,dt}, 6, 12);

                }
                catch (Exception)
                {
                }
            }
            
        }
    }

    public abstract class BaoCaoHCMExportDiemMonhoc
    {
        public BaoCaoHCMExportDiemMonhoc(string tenmonhoc, string namhoc)
        {
            this.NamHoc = namhoc;
            this.TenMonHoc = tenmonhoc;
        }
        protected string TenMonHoc { get; set; }
        protected string NamHoc { get; set; }
        protected string TieudeThiHK
        {
            get
            {
                return string.Format("ĐIỂM SỐ MÔN {0} KIỂM TRA HKII  NĂM HỌC {1}", this.TenMonHoc, this.NamHoc);
            }
        }
        protected string TieudeXLHK2
        {
            get
            {
                return string.Format("XẾP LOẠI  MÔN {0} HK II -  NĂM HỌC {1}", this.TenMonHoc, this.NamHoc);
            }
        }

        protected string TieudeXLCN
        {
            get { return string.Format("XẾP LOẠI  MÔN {0} - CẢ NĂM -  NĂM HỌC {1}", this.TenMonHoc, this.NamHoc); }
        }
        public abstract void Export(string fileTemplate, string name, DataTable[] dataSource, int hangBatDauBangChiTiet, int hangBatDauBangToanTruong);

        protected string GetAddressString(IXLCell cellAddress, bool type = false)
        {
            if (type)
            {
                // có $
                return string.Format("{0}${1}", cellAddress.Address.ColumnLetter, cellAddress.Address.RowNumber);
            }
            else
            {
                // ko có $
                return string.Format("{0}", cellAddress.Address.ToString());
            }
        }

        protected void SetStylePage(IXLWorksheet worksheet)
        {

            worksheet.PageSetup.Margins.Left = 0.7;
            worksheet.PageSetup.Margins.Right = 0.7;
            worksheet.PageSetup.Margins.Top = 0.5;
            worksheet.PageSetup.Margins.Bottom = 0;
            worksheet.PageSetup.Margins.Header = 0.5;
            worksheet.PageSetup.Margins.Footer = 0.5;
        }

        protected void Response(XLWorkbook workbook,string name)
        {
            HttpResponse _response = HttpContext.Current.Response;
            _response.ClearContent();
            _response.Buffer = true;
            _response.AddHeader("content-disposition", string.Format("attachment; filename={0}.xlsx", name ?? (new Guid()).ToString()));
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

        protected void SetTitle(IXLWorksheet sheet)
        {
            
        }
        
    }

    public class MonHoc
	{
        public string MaMonHoc { get; set; }
        public string TenMonHoc { get; set; }
	}

    /// <summary>
    /// Báo cáo Hồ Chí Minh Môn điểm số
    /// </summary>
    public class BaoCaoHCM_C2RenderMonDiemSo : BaoCaoHCMExportDiemMonhoc
    {
        public BaoCaoHCM_C2RenderMonDiemSo(string tenmonhoc, string namhoc) : base(tenmonhoc, namhoc) { }
        public override void Export(string fileTemplate, string name,DataTable[] dataSource, int hangBatDauBangChiTiet, int hangBatDauBangToanTruong)
        {
            try
            {
                using (XLWorkbook workbook = new XLWorkbook(fileTemplate))
                {
                    if (dataSource.Length == 3)
                    {
                        #region lấy danh sách worksheet và thiết lập tiêu đề môn học
                        
                        var sheet = workbook.Worksheet(1);
                        sheet.Cell(1, 5).Value = this.TieudeThiHK;
                        var sheet2 = workbook.Worksheet(2);
                        sheet.Cell(1, 5).Value = this.TieudeXLHK2;
                        var sheet3 = workbook.Worksheet(3);
                        sheet.Cell(1, 5).Value = this.TieudeXLCN;

                        #endregion


                        #region Thiết lập danh sách cột dữ liệu theo các trường được trả về từ store procedures
                        
                        List<ColumnDetail> columns = new List<ColumnDetail>();
                        columns.Add(new ColumnDetail { FieldName = "TEN_TRUONG" });
                        columns.Add(new ColumnDetail { FieldName = "SO_LUONGHS" });
                        columns.Add(new ColumnDetail { FieldName = "GIOI_TS" });
                        columns.Add(new ColumnDetail { FieldName = "GIOI_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "KHA_TS" });
                        columns.Add(new ColumnDetail { FieldName = "KHA_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "TRUNGBINH_TS" });
                        columns.Add(new ColumnDetail { FieldName = "TRUNGBINH_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "YEU_TS" });
                        columns.Add(new ColumnDetail { FieldName = "YEU_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "KEM_TS" });
                        columns.Add(new ColumnDetail { FieldName = "KEM_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "TBTROLEN_TS" });
                        columns.Add(new ColumnDetail { FieldName = "TBTROLEN_TL", DataType = XLCellValues.Number });
                        columns.Add(new ColumnDetail { FieldName = "TBTROLEN_XEP_HANG" });

                        #endregion

                        ////Xuất dữ liệu cho các sheet tướng ứng
                        this.GetWorkSheet(dataSource[0], sheet, columns, hangBatDauBangChiTiet, hangBatDauBangToanTruong);
                        this.GetWorkSheet(dataSource[1], sheet2, columns, hangBatDauBangChiTiet, hangBatDauBangToanTruong);
                        this.GetWorkSheet(dataSource[2], sheet3, columns, hangBatDauBangChiTiet, hangBatDauBangToanTruong);
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
            sheet.Column(1).Width = 18;
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
        private List<int> SetToanTruong(IXLWorksheet sheet, List<int> indexs, ref int rowStart, List<ColumnDetail> columns)
        {
            List<int> range = new List<int>();
            sheet.Row(rowStart).InsertRowsBelow(indexs.Count + 1);
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

                        case "TEN_TRUONG":
                            cell.Value = khoi[j];
                            break;
                        case "SO_LUONGHS":
                            cell.FormulaA1 = string.Format("=SUM({0},{1},{2},{3},{4})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(rowStart, 7).Address.ToString(), sheet.Cell(rowStart, 9).Address.ToString(), sheet.Cell(rowStart, 11).Address.ToString());
                            break;
                        case "GIOI_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 3).Address.ToString());
                            break;
                        case "GIOI_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            break;
                        case "KHA_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 5).Address.ToString());
                            break;
                        case "KHA_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            break;
                        case "TRUNGBINH_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 7).Address.ToString());
                            break;
                        case "TRUNGBINH_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            break;
                        case "YEU_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 9).Address.ToString());
                            break;
                        case "YEU_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            break;
                        case "KEM_TS":
                            cell.FormulaA1 = string.Format("={0}", sheet.Cell(indexs[j], 11).Address.ToString());
                            break;
                        case "KEM_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            break;
                        case "TBTROLEN_TS":
                            cell.FormulaA1 = string.Format("=SUM({0},{1},{2})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(rowStart, 7).Address.ToString());
                            break;
                        case "TBTROLEN_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 13).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                            break;
                        case "TBTROLEN_XEP_HANG":
                            cell.FormulaA1 = cell.FormulaA1 = string.Format("=RANK({0},{1}:{2},0)", sheet.Cell(rowStart, 14).Address.ToString(), this.GetAddressString(sheet.Cell(rowBegin, 14), true), this.GetAddressString(sheet.Cell(rowBegin + khoi.Length - 1, 14), true));
                            break;
                        default:
                            break;
                    }
                    #endregion

                    //set cell type
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
            List<string> cellsGioiTS = new List<string>();
            List<string> cellsKhaTS = new List<string>();
            List<string> cellsTBTS = new List<string>();
            List<string> cellsYeuTS = new List<string>();
            List<string> cellsKemTS = new List<string>();
            List<string> cellsTBTroLenTS = new List<string>();
            foreach (var item in indexs)
            {
                cellsGioiTS.Add(sheet.Cell(item, 3).Address.ToString());
                cellsKhaTS.Add(sheet.Cell(item, 5).Address.ToString());
                cellsTBTS.Add(sheet.Cell(item, 7).Address.ToString());
                cellsYeuTS.Add(sheet.Cell(item, 9).Address.ToString());
                cellsKemTS.Add(sheet.Cell(item, 11).Address.ToString());
                cellsTBTroLenTS.Add(sheet.Cell(item, 13).Address.ToString());
            }
            for (int i = 1; i <= columns.Count; i++)
            {
                var cell = sheet.Cell(rowStart, i);
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                #region Thiết lập cột
                switch (columns[i - 1].FieldName)
                {
                    case "TEN_TRUONG":
                        cell.Value = "THCS";
                        break;
                    case "SO_LUONGHS":
                        cell.FormulaA1 = string.Format("=SUM({0},{1},{2},{3},{4})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(rowStart, 7).Address.ToString(), sheet.Cell(rowStart, 9).Address.ToString(), sheet.Cell(rowStart, 11).Address.ToString());
                        break;
                    case "GIOI_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsGioiTS));
                        break;
                    case "GIOI_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        break;
                    case "KHA_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsKhaTS));
                        break;
                    case "KHA_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        break;
                    case "TRUNGBINH_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsTBTS));
                        break;
                    case "TRUNGBINH_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 7).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        break;
                    case "YEU_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsYeuTS));
                        break;
                    case "YEU_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 9).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        break;
                    case "KEM_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsKemTS));
                        break;
                    case "KEM_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 11).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        break;
                    case "TBTROLEN_TS":
                        cell.FormulaA1 = string.Format("={0}", string.Join("+", cellsTBTroLenTS));
                        break;
                    case "TBTROLEN_TL":
                        cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(rowStart, 13).Address.ToString(), sheet.Cell(rowStart, 2).Address.ToString());
                        break;
                    case "TBTROLEN_XEP_HANG":

                        break;
                    default:
                        break;

                }
                #endregion
                if (columns[i - 1].DataType== XLCellValues.Number)
                {
                    cell.DataType = XLCellValues.Number;
                    cell.Style.NumberFormat.NumberFormatId = 10;
                }
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
                    if (i == 1)
                    {
                        //tên trường
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    }
                    else //cột thường
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    #region Thiết lập cột
                    switch (col.FieldName)
                    {

                        case "TEN_TRUONG": cell.Value = item[col.FieldName];
                            if (ro == dt.Rows.Count - 1)
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            break;
                        case "SO_LUONGHS":
                            cell.FormulaA1 = string.Format("=SUM({0},{1},{2},{3},{4})", sheet.Cell(row, 3).Address.ToString(), sheet.Cell(row, 5).Address.ToString(), sheet.Cell(row, 7).Address.ToString(), sheet.Cell(row, 9).Address.ToString(), sheet.Cell(row, 11).Address.ToString());
                            break;
                        case "GIOI_TS":
                            if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 3).Address.ToString(), sheet.Cell(row - 1, 3).Address.ToString());
                            else
                        cell.Value = item[col.FieldName];
                            break;
                        case "GIOI_TL": 
                            cell.FormulaA1= string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 3).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "KHA_TS":
                            if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 5).Address.ToString(), sheet.Cell(row - 1, 5).Address.ToString());
                            else
                                cell.Value = item[col.FieldName];
                            break;
                        case "KHA_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 5).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "TRUNGBINH_TS":
                            if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 7).Address.ToString(), sheet.Cell(row - 1, 7).Address.ToString());
                            else
                                cell.Value = item[col.FieldName];
                            break;
                        case "TRUNGBINH_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 7).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "YEU_TS":
                            if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 9).Address.ToString(), sheet.Cell(row - 1, 9).Address.ToString());
                            else
                                cell.Value = item[col.FieldName];
                            break;
                        case "YEU_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 9).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "KEM_TS":
                            if (ro == dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=SUM({0}:{1})", sheet.Cell(rowStart, 11).Address.ToString(), sheet.Cell(row - 1, 11).Address.ToString());
                            else
                                cell.Value = item[col.FieldName];
                            break;
                        case "KEM_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 11).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "TBTROLEN_TS":
                            cell.FormulaA1 = string.Format("=SUM({0},{1},{2})", sheet.Cell(row, 3).Address.ToString(), sheet.Cell(row, 5).Address.ToString(), sheet.Cell(row, 7).Address.ToString());
                            break;
                        case "TBTROLEN_TL":
                            cell.FormulaA1 = string.Format("=IF({0},{0}/${1},0)", sheet.Cell(row, 13).Address.ToString(), sheet.Cell(row, 2).Address.ToString());
                            break;
                        case "TBTROLEN_XEP_HANG":
                            if (ro != dt.Rows.Count - 1)
                                cell.FormulaA1 = string.Format("=RANK({0},{1}:{2},0)", sheet.Cell(row, 14).Address.ToString(), this.GetAddressString(sheet.Cell(rowStart, 14), true), this.GetAddressString(sheet.Cell(rowStart + dt.Rows.Count - 2, 14), true));
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

        protected string GetAddressString(IXLCell cellAddress, bool type=false)
        {
            if (type)
            {
                // có $
                return string.Format("{0}${1}", cellAddress.Address.ColumnLetter, cellAddress.Address.RowNumber);
            }
            else
            {
                // ko có $
                return string.Format("{0}", cellAddress.Address.ToString());
            }
        }
    }
}