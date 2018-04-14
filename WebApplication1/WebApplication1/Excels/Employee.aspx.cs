using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebApplication1.Design;
using WebApplication1.ExcelCore;
using WebApplication1.Models;
using WebApplication1.Models.Entites;
using WebApplication1.Models.SqlHelper;

namespace WebApplication1.Excels
{
    public partial class Employee : System.Web.UI.Page
    {
        protected List<NhanVien> employees;

        protected bool isDemo;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!this.IsPostBack)
            {
                LoadChiNhanh();
                ddlChiNhanh.SelectedIndex = 0;
                ddlPhongBan.Items.Clear();
                ddlPhongBan.Items.Add(new ListItem { Value = "All", Text = "Tất cả", Selected = true });
            }
            this.employees = new List<NhanVien>();
        }

        protected void btnChange_Click(object sender, EventArgs e)
        {
            DataTable dt = LoadEmployee(ddlChiNhanh.SelectedValue, ddlPhongBan.SelectedValue);
            ckblColumns.Items.Clear();
            for (int i = 0; i < cols.Length; i++)
            {
                ckblColumns.Items.Add(new ListItem(colsname[i], cols[i]));
            }
        }
        private string[] cols = new string[] { "MaNV", "TenNV", "MaPB", "NgaySinh", "DiaChi", "DienThoai", "CMND", "GioiTinh", "Diem", "TenPB" };
        private string[] colsname = new string[] { "Mã", "Họ tên", "Mã phòng ban", "Ngày sinh", "Địa chỉ", "Điện thoại", "CMND", "Giới tính", "Điểm", "Tên phòng ban" };
        private void LoadChiNhanh()
        {
            DataTable dt = new Chinhanh().GetList();
            ddlChiNhanh.Items.Clear();
            ddlChiNhanh.Items.Add(new ListItem { Value = "All", Text = "Tất cả", Selected = true });
            foreach (DataRow item in dt.Rows)
            {
                ddlChiNhanh.Items.Add(new ListItem { Value = item["MaChiNhanh"].ToString(), Text = item["TenChiNhanh"].ToString() });
            }
        }

        private DataTable LoadEmployee(string macn, string mapb)
        {
            DataTable dt = this.GetData(macn, mapb);
            foreach (DataRow item in dt.Rows)
            {
                NhanVien nv = new NhanVien();
                nv.MaNV = item["MaNV"].ToString();
                nv.TenNV = item["TenNV"].ToString();
                nv.MaPB = item["MaPB"].ToString();
                nv.NgaySinh = Convert.ToDateTime(item["NgaySinh"].ToString());
                nv.DienThoai = item["DienThoai"].ToString();
                nv.DiaChi = item["DiaChi"].ToString();
                employees.Add(nv);
            }
            return dt;
        }

        private void LoadPhongBan(string machinhanh)
        {
            ddlPhongBan.Items.Clear();
            ddlPhongBan.Items.Add(new ListItem { Value = "All", Text = "Tất cả", Selected = true });
            DataTable dt = new Phongban().GetList(machinhanh);
            foreach (DataRow item in dt.Rows)
            {
                ddlPhongBan.Items.Add(new ListItem { Value = item["MaPB"].ToString(), Text = item["TenPB"].ToString() });
            }
        }

        protected void ddlChiNhanh_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadPhongBan(ddlChiNhanh.SelectedValue);
        }

        protected void btnExport2_Click(object sender, EventArgs e)
        {
            BookManager manager = new BookManager("Danh sách nhân viên ");
            var d = ddlPhongBan.Items;
            string mapb = ddlPhongBan.SelectedValue;
            string macn = ddlChiNhanh.SelectedValue;
            WorkSheetSetting setting = new WorkSheetSetting();
            List<ExcelCore.ExcelColumnContent> colmuns = new List<ExcelCore.ExcelColumnContent> {
                    new ExcelColumnContent{ Width=15, Name="Mã số"},
                    new ExcelColumnContent{ Width=25, Name="Tên nhân viên"},
                    new ExcelColumnContent{ Width=18,Name="Ngày sinh"},
                    new ExcelColumnContent{ Width=25,Name="Địa chỉ"},
                    new ExcelColumnContent{ Width=16,Name="Số điện thoại"},
                    new ExcelColumnContent{ Width=15,Name="CMND"},
                    new ExcelColumnContent{ Width=8,Name="Điểm"}
                    };
            setting.BoderStyle = XLBorderStyleValues.Thin;
            setting.ShowGridLine = false;
            setting.WrapText = false;
            setting.TableHeaderBold = true;
            setting.PageOrientation = XLPageOrientation.Landscape;
            setting.ColumnsWidth = colmuns.ToArray();
            if (macn == "All")
            {
                //dt = new DataTable();
            }
            else
            {
                List<DataTable> tables = new List<DataTable>();
                if (mapb == "All")
                {
                    for (int i = 1; i < d.Count; i++)
                    {
                        DataTable dt = new NhanVien().GetByPhongban(d[i].Value);
                        dt.TableName = d[i].Text;
                        tables.Add(dt);
                    }
                }
                else
                {
                    DataTable dt = new NhanVien().GetByPhongban(mapb);
                    dt.TableName = ddlPhongBan.SelectedItem.Text;
                    tables.Add(dt);
                }
                int j = 1;
                foreach (var item in tables)
                {
                    item.Columns.Remove(item.Columns["MaPB"]);
                    item.Columns.Remove(item.Columns["TenPB"]);
                    item.Columns.Remove(item.Columns["GioiTinh"]);
                    WorkSheetSetting set = setting.Clone();
                    set.SheetName = item.TableName;
                    set.TitleSheet = "Danh sách nhân viên " + item.TableName;
                    set.DataSource = item;
                    //if (j % 2 == 0)
                    //manager.AddSheet(new WorksheetTemplateBase(set));
                    //else 
                    manager.AddSheet(new WorksheetReportColor(set));
                    j++;
                }
                manager.Export();
            }
        }

        private DataTable GetData(string macn, string mapb)
        {
            DataTable dt;
            if (macn == "All")
            {
                //lấy tất cả các chi nhánh
                dt = new NhanVien().GetList();
            }
            else
            {
                if (mapb == "All")
                {
                    //lấy tất cả các phòng trong chi nhánh
                    dt = new NhanVien().GetByChinhanh(macn);
                }
                else
                {
                    //lấy theo chi nhánh
                    dt = new NhanVien().GetByPhongban(mapb);
                }
            }
            return dt;
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            string macn = ddlChiNhanh.SelectedValue;
            string mapb = ddlPhongBan.SelectedValue;
            DataTable dsnhanvien = this.GetData(macn, mapb);
            if (ddlPhongBan.SelectedItem != null)
            {
                string tenpb = ddlPhongBan.SelectedItem == null ? "Tiêu đề demo" : ddlPhongBan.SelectedItem.Text;
                BookManager manager = new BookManager("Danh sách nhân viên " + tenpb);
                WorkSheetSetting setting = new WorkSheetSetting();
                dsnhanvien.Columns.Remove(dsnhanvien.Columns["MaPB"]);
                dsnhanvien.Columns.Remove(dsnhanvien.Columns["GioiTinh"]);
                dsnhanvien.Columns.Remove(dsnhanvien.Columns["CMND"]);
                List<ExcelCore.ExcelColumnContent> colmuns = new List<ExcelCore.ExcelColumnContent> {
                    new ExcelColumnContent{ Width=12, Name="Mã số"},
                    new ExcelColumnContent{ Width=23, Name="Tên nhân viên"},
                    new ExcelColumnContent{ Width=18,Name="Ngày sinh"},
                    new ExcelColumnContent{ Width=16,Name="Địa chỉ"},
                    new ExcelColumnContent{ Width=16,Name="Số điện thoại"},
                    new ExcelColumnContent{ Width=6,Name="Điểm"}
                    };
                if (macn == "All")
                {
                    colmuns.Add(new ExcelColumnContent { Width = 10, Name = "Phòng ban" });
                    colmuns.Add(new ExcelColumnContent { Width = 10, Name = "Chi nhánh" });
                }
                else
                {
                    if (mapb == "All")
                    {
                        colmuns.Add(new ExcelColumnContent { Width = 20, Name = "Phòng ban" });
                    }
                    else
                    {
                        dsnhanvien.Columns.Remove(dsnhanvien.Columns["TenPB"]);
                    }
                }
                setting.BoderStyle = XLBorderStyleValues.Thin;
                setting.SheetName = tenpb;
                setting.ShowGridLine = false;
                setting.TitleSheet = "Danh sách nhân viên " + tenpb;
                setting.WrapText = false;
                setting.DataSource = dsnhanvien;
                setting.TableHeaderBold = true;
                setting.PageOrientation = XLPageOrientation.Landscape;
                setting.ColumnsWidth = colmuns.ToArray();
                manager.AddSheet(new WorksheetTemplateBase(setting));
                manager.Export();
            }
        }


        protected void btnAdvanted_Click(object sender, EventArgs e)
        {
            BookDemo manager = new BookDemo("Danh sách nhân viên ");
            var d = ddlPhongBan.Items;
            string mapb = ddlPhongBan.SelectedValue;
            string macn = ddlChiNhanh.SelectedValue;
            //WorksheetConfig setting = new WorksheetConfig();
            List<ExcelCore.ExcelColumnContent> columns = new List<ExcelColumnContent>();
            for (int i = 0; i < ckblColumns.Items.Count; i++)
            {
                if (ckblColumns.Items[i].Selected)
                {
                    columns.Add(new ExcelColumnContent { Name = ckblColumns.Items[i].Value, Caption = ckblColumns.Items[i].Text });
                }
            }
            //List<ExcelCore.ExcelColumnContent> columns = new List<ExcelCore.ExcelColumnContent> {
            //        new ExcelColumnContent{ Width=15, Caption="Mã số",Name="MaNV"},
            //        new ExcelColumnContent{ Width=25, Caption="Tên nhân viên",Name="TenNV"},
            //        new ExcelColumnContent{ Width=18,Caption="Ngày sinh",Name="NgaySinh"},
            //        new ExcelColumnContent{ Width=25,Caption="Địa chỉ",Name="DiaChi"},
            //        new ExcelColumnContent{ Width=16,Caption="Số điện thoại",Name="DienThoai"},
            //        new ExcelColumnContent{ Width=15,Caption="CMND",Name="CMND"},
            //        new ExcelColumnContent{ Width=8,Caption="Điểm",Name="Diem"}
            //        };
            


            

            //setting.Title.SubTitle.Text = string.Format("Từ ngày {0} đến {1}", DateTime.Now.AddDays(-30).ToString("dd/MM/yyyy"), DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"));
            
            //setting.Header.HeaderLeft.Add(new TextComponent { Text = "SỞ Y TẾ HÀ NỘI", Style = styleheader });
            //setting.Header.HeaderLeft.Add(new TextComponent { Text = "BỆNH VIỆN ĐA KHOA ĐAN PHƯỢNG", Style = styleHeaderBold });
            //setting.Header.HeaderRight.Add(new TextComponent { Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", Style = styleHeaderBold });
            //setting.Header.HeaderRight.Add(new TextComponent { Text = "Độc lập - Tự do - Hạnh phúc", Style = styleHeaderBold });
            //setting.Title.Title.Style = styletitle;
            if (mapb == "All")
            {
                for (int i = 1; i < d.Count; i++)
                {
                    //WorksheetConfig set = ExcelCommon.DeepCopy<WorksheetConfig>(setting);
                    WorksheetLayoutLevel1 layout1 = new WorksheetLayoutLevel1();
                    DataTable dt = new NhanVien().GetByPhongban(d[i].Value);
                    columns = ExcelCommon.GetColumnsValid(dt, columns);
                    //columns.ForEach(p => p.Width = 128 / columns.Count);
                    DataTable dnew = dt.DefaultView.ToTable(true, columns.Select(p => p.Name).ToArray());
                    layout1.SetName ( d[i].Text);
                    layout1.SetTitle("Danh sách nhân viên " + d[i].Text);
                    if (i!=1)
                    {
                        layout1.SetTitle("Danh sách nhân viên 1 " + d[i].Text);
                        layout1.Setting.Title.Title.Style.FontSize = 33;
                    }
                    layout1.SetDataSource( dnew);
                    layout1.SetColumns(columns);
                    manager.AddSheet(layout1);
                }
            }
            else
            {
                //WorksheetConfig set = ExcelCommon.DeepCopy<WorksheetConfig>(setting);
                WorksheetLayoutLevel1 layout1 = new WorksheetLayoutLevel1();
                DataTable dt = new NhanVien().GetByPhongban(mapb);
                columns = ExcelCommon.GetColumnsValid(dt, columns);
                //columns.ForEach(p => p.Width = 128 / columns.Count);
                DataTable dnew = dt.DefaultView.ToTable(true, columns.Select(p => p.Name).ToArray());
                layout1.SetColumns(columns);
                layout1.SetName(ddlPhongBan.SelectedItem.Text);
                layout1.SetTitle("Danh sách nhân viên " + ddlPhongBan.SelectedItem.Text);
                layout1.SetDataSource(dnew);
                manager.AddSheet(layout1);
            }
            manager.Export();
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            using (XLWorkbook workbook=new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("demo");
                worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
                worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
                worksheet.PageSetup.AdjustTo(100);
                worksheet.PageSetup.Margins.Left = 0.7;
                worksheet.PageSetup.Margins.Right = 0.7;
                worksheet.PageSetup.Margins.Top = 0.5;
                worksheet.PageSetup.Margins.Bottom = 0;
                worksheet.PageSetup.Margins.Header = 0.5;
                worksheet.PageSetup.Margins.Footer = 0.5;
                worksheet.Cell(1, 1).Value = "Nguyễn Đình Khánh";
                worksheet.Cell(1, 2).Value = "Nguyễn Đình Khánh 1";
                worksheet.Cell(1, 3).Value = "Nguyễn Đình Khánh 2 ";
                //worksheet.Column(1).Width = 25;
                //worksheet.Column(2).Width = 10;
                //worksheet.Column(3).Width = 10;
                worksheet.Cell(1, 1).Style.Border.TopBorder =  XLBorderStyleValues.Thin;
                worksheet.Cell(1, 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 1).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 2).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 2).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 2).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 2).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 3).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 3).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 3).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(1, 3).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                HttpResponse _response = HttpContext.Current.Response;
                _response.ClearContent();
                _response.Buffer = true;
                _response.AddHeader("content-disposition", string.Format("attachment; filename={0}.xlsx","filename" ?? (new Guid()).ToString()));
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
        }
    }
}