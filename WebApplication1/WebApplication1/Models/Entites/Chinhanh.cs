using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using WebApplication1.Models.SqlHelper;

namespace WebApplication1.Models.Entites
{
    public class Chinhanh : ISqlAction<Chinhanh,string>
    {
        public string MaChiNhanh { get; set; }
        public string TenChiNhanh { get; set; }
        public string DiaChi { get; set; }
        public string DienThoai { get; set; }
        public string Email { get; set; }

        public void Insert(Chinhanh item)
        {
            throw new NotImplementedException();
        }

        public void Insert()
        {
            throw new NotImplementedException();
        }

        public System.Data.DataTable GetList()
        {
            SqlAccess access = new SqlAccess();
            DataTable dt = access.ExecuteStore("sp_ChinhanhGetAll");
            return dt;
        }

        public string Update(Chinhanh item)
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

        public string Delete(Chinhanh item)
        {
            throw new NotImplementedException();
        }

        public Chinhanh GetItem(string key)
        {
            throw new NotImplementedException();
        }

        public Chinhanh GetItem()
        {
            throw new NotImplementedException();
        }
    }
}