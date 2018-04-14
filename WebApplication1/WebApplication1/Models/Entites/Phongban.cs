using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using WebApplication1.Models.SqlHelper;

namespace WebApplication1.Models.Entites
{
    public class Phongban : ISqlAction<Phongban,string>
    {
        public string MaPB { get; set; }
        public string TenPB { get; set; }
        public string MaChiNhanh { get; set; }
        public string DienThoai { get; set; }

        public void Insert(Phongban item)
        {
            throw new NotImplementedException();
        }

        public void Insert()
        {
            throw new NotImplementedException();
        }
        public System.Data.DataTable GetList(string macn)
        {
            SqlAccess access = new SqlAccess();
            DataTable dt = access.ExecuteStore("sp_PhongbanGetChinhanh", new string[] { "@MaCN" }, new object[] { macn });
            return dt;
        }


        public System.Data.DataTable GetList()
        {
            throw new NotImplementedException();
        }

        public string Update(Phongban item)
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

        public string Delete(Phongban item)
        {
            throw new NotImplementedException();
        }

        public Phongban GetItem(string key)
        {
            throw new NotImplementedException();
        }

        public Phongban GetItem()
        {
            throw new NotImplementedException();
        }
    }
}