using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using XuatExcelClosedXML.Models.SqlHelper;

namespace XuatExcelClosedXML.Providers
{
    public class BaoCaoHCM_C2
    {
        protected string constr = System.Configuration.ConfigurationManager.ConnectionStrings["HCM_HTTT_C2"].ConnectionString;
        protected SqlAccess sqlAccess;
        public BaoCaoHCM_C2()
        {
            this.sqlAccess = new SqlAccess(this.constr);
        }
        public DataTable GetDiemSoMonHoc(string mapgd, string manamhoc, string mamonhoc, int hocky, int solieu)
        {
            DataTable data = this.sqlAccess.ExecuteStore("BaoCao_DiemSo_XepLoai_MonHoc_CapPhong",
                new string[] { "@MaPGD", "@MaNamHoc", "@MaMonHoc", "@HocKy", "@SoLieu" },
                new object[] { mapgd, manamhoc, mamonhoc, hocky, solieu });
            return data;
        }
    }
}