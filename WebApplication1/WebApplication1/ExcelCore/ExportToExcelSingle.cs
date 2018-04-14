using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication1.ExcelCore
{
    public class ExportToExcelSingle
    {
        private DataTable data;
        private XLWorkbook workbook;
        private string _filename;
        private string _title;
        public ExportToExcelSingle(DataTable dt,string name)
        {
            this.data = dt;
            this._title = name;
            this._filename = name;
            this.workbook = new XLWorkbook();
        }

        protected void CreateSheet()
        {
                var sheet = this.workbook.AddWorksheet(this._title);
                ExcelSimple simple = new ExcelSimple(sheet);

                simple.BoderStyle = XLBorderStyleValues.Thin;
                simple.SheetsName = "Nhan vien";
                simple.ShowGridLine = false;
                simple.TitleSheet = "Danh sách nhân viên " + this._title;
                simple.WrapText = false;
                simple.DataSource = this.data;
                simple.TableHeaderBold = true;
                simple.PageOrientation = XLPageOrientation.Landscape;
                simple.ColumnsWidth = new ExcelCore.ExcelColumnContent[9] {
                            new ExcelColumnContent{ Width=12, Name="Mã số"},
                            new ExcelColumnContent{ Width=20, Name="Tên nhân viên"},
                            new ExcelColumnContent{ Width=10,Name="Phòng ban", HorizontalAlignment= XLAlignmentHorizontalValues.Center},
                            new ExcelColumnContent{ Width=18,Name="Ngày sinh"},
                            new ExcelColumnContent{ Width=15,Name="Địa chỉ"},
                            new ExcelColumnContent{ Width=16,Name="Số điện thoại"},
                            new ExcelColumnContent{ Width=15,Name="CMND"},
                            new ExcelColumnContent{ Width=8,Name="Giới tính"},
                            new ExcelColumnContent{ Width=6,Name="Điểm"},
                            };
                sheet = simple.Init();
            
        }
        public void Export()
        {
            this.CreateSheet();
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
    }
}