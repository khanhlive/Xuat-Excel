using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using XuatExcelClosedXML.Models.SqlHelper;

namespace XuatExcelClosedXML.Models.Entites
{
    public class NhanVien : ISqlAction<NhanVien,string>
    {
        public NhanVien() { }
        public NhanVien(string manv, string tennv, string mapb, DateTime ngaysinh, string diachi, string dienthoai,string cmnd,bool gioitinh,double diem)
        {
            this.MaNV = manv;
            TenNV = tennv;
            MaPB = mapb;
            NgaySinh = ngaysinh;
            DiaChi = diachi;
            DienThoai = dienthoai;
            this.CMND = cmnd;
            this.Diem = diem;
            this.GioiTinh = gioitinh;
        }
        public string MaNV { get; set; }
        public string TenNV { get; set; }
        public string MaPB { get; set; }
        public DateTime NgaySinh { get; set; }
        public string DiaChi { get; set; }
        public string DienThoai { get; set; }
        public string CMND { get; set; }
        public bool GioiTinh { get; set; }
        public double Diem { get; set; }

        public void Insert(NhanVien item)
        {
            throw new NotImplementedException();
        }

        public void Insert()
        {
            throw new NotImplementedException();
        }

        public DataTable GetByPhongban(string mapb)
        {
            SqlAccess access = new SqlAccess();
            DataTable dt = access.ExecuteStore("sp_NhanvienGetPhongban", new string[] { "@MaPB" }, new object[] { mapb });
            return dt;
        }
        public DataTable GetByChinhanh(string macn)
        {
            SqlAccess access = new SqlAccess();
            DataTable dt = access.ExecuteStore("sp_NhanvienGetChinhanh", new string[] { "@MaCN" }, new object[] { macn });
            return dt;
        }

        public DataTable GetList()
        {
            SqlAccess access = new SqlAccess();
            DataTable dt = access.ExecuteStore("sp_NhanvienGetAll");
            return dt;
        }

        public string Update(NhanVien item)
        {
            throw new NotImplementedException();
        }

        public string Update()
        {
            throw new NotImplementedException();
        }

        public string Delete()
        {
            throw new NotImplementedException();
        }

        public string Delete(NhanVien item)
        {
            throw new NotImplementedException();
        }

        public NhanVien GetItem(string key)
        {
            throw new NotImplementedException();
        }

        public NhanVien GetItem()
        {
            throw new NotImplementedException();
        }
    }
}