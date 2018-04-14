using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using WebApplication1.Providers;

namespace WebApplication1.Excels
{
    public partial class BaocaoHCM : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        protected void btnExport_Click(object sender, EventArgs e)
        {
            if (fileUpload.HasFile)
            {
                string filePathServer="";
                try
                {
                    string fileName = Path.GetFileName(fileUpload.FileName);
                    filePathServer = Server.MapPath("~/Uploads/" + fileName);
                    fileUpload.SaveAs(filePathServer);

                    string MaPGD = "hcm-q1",
                        MaNamHoc = "2016",
                        MaMonHoc = "MT";
                    int NamHocID = 9551, HocKy = 2, KieuMon = 2, SoLieu = 1;
                    BaoCaoHCM_C2 baocaohcm = new BaoCaoHCM_C2();
                    //DataTable tableThiHK = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamHoc, MaMonHoc, NamHocID, HocKy, KieuMon, SoLieu);
                    //SoLieu = 2;
                    //DataTable tableXLHK = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamHoc, MaMonHoc, NamHocID, HocKy, KieuMon, SoLieu);
                    //HocKy = 3;
                    //DataTable tableXLCN = baocaohcm.GetDiemSoMonHoc(MaPGD, MaNamHoc, MaMonHoc, NamHocID, HocKy, KieuMon, SoLieu);
                    //List<DataColumn> columns = new List<ColumnDetail>();

                    #region Tạo data mẫu

                    DataTable dt = new DataTable();
                    dt.Columns.Add("KHOI" );
                    dt.Columns.Add("TEN_TRUONG");
                    dt.Columns.Add("SO_LUONGHS");
                    dt.Columns.Add("GIOI_TS");
                    dt.Columns.Add("GIOI_TL");
                    dt.Columns.Add("KHA_TS");
                    dt.Columns.Add("KHA_TL");
                    dt.Columns.Add("TRUNGBINH_TS");
                    dt.Columns.Add("TRUNGBINH_TL");
                    dt.Columns.Add("YEU_TS");
                    dt.Columns.Add("YEU_TL");
                    dt.Columns.Add("KEM_TS");
                    dt.Columns.Add("KEM_TL");
                    dt.Columns.Add("TBTROLEN_TS");
                    dt.Columns.Add("TBTROLEN_TL");
                    dt.Columns.Add("TBTROLEN_XEP_HANG");
                    string[] arrTruong = new string[] { "Trường 1", "Trường 2", "Trường 3", "Trường 4", "Trường 5", "Trường 6", "Trường 7", "Trường 8" };
                    int khoi = 6;
                    for (int i = 1; i <= 36; i++)
                    {
                        if (i % arrTruong.Length == 0)
                        {
                            dt.Rows.Add(new string[] { (khoi*10).ToString(),"Khối " + khoi.ToString() });
                            //dr[0] = "Khối " + khoi.ToString();
                            khoi++;
                            //dt.Rows.Add(dr);
                        }
                        else
                        {
                            List<object> row = new List<object>();
                            row.Add(khoi);
                            row.Add(arrTruong[new Random().Next(arrTruong.Length-1)]);
                            for (int j = 2; j < dt.Columns.Count; j++)
                            {
                                row.Add(new Random().Next(1000));
                            }
                            dt.Rows.Add(row.ToArray());
                        }
                    }
                    #endregion

                    BaoCaoHCMManager.XuatBaoCaoHCM(filePathServer,"Môn Tiếng Anh","2017 - 2018", new DataTable[] { dt, dt, dt }, 6, 12);
                    //BaoCaoHCMManager.XuatBaocaoHCM_Nhieumon(filenameServer, filenameServer, new List<MonHoc> { new MonHoc { MaMonHoc = "12", TenMonHoc = "ten 1" }, new MonHoc { MaMonHoc = "122", TenMonHoc = "ten 2" } }, dt, "", "", 0, 0);
                    //BaoCaoHCMManager.ExportBaoCaoHCM(filenameServer, new DataTable[] { tableThiHK, tableXLHK, tableXLCN }, 5, 62);
                    
                }
                catch (Exception ex)
                {
                    Common.log.Error("Error", ex);
                }
                finally
                {
                    if (System.IO.File.Exists(filePathServer))
                        System.IO.File.Delete(filePathServer);
                }
            }
                
        }
    }
}