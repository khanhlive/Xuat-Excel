using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace XuatExcelClosedXML.Models.SqlHelper
{
    public class SqlAccess
    {
        protected string conStr = ConfigurationManager.ConnectionStrings["ExcelExampleConnection"].ConnectionString;
        protected SqlConnection connection;
        protected SqlCommand cmd;
        public SqlAccess()
        {
            this.connection = new SqlConnection(this.conStr);
            this.cmd = null;
        }
        public SqlAccess(string _constr)
        {
            //this.conStr = this.conStr;
            this.connection = new SqlConnection(_constr);
            this.cmd = null;
        }

        protected SqlConnection GetConnection()
        {
            if (this.connection==null)
            {
                this.connection = new SqlConnection(this.conStr);
            }
            return this.connection;
        }

        public void InitData()
        {
            string[] names = new string[] {"Nguyễn" ,"Phạm","Trần","Lê","Đỗ"};
            string[] name2=new string[]{"Văn","Đình","Bá","Quang","Việt","Minh"};
            string[] name3 = new string[] { "Khánh", "Hải", "Sơn", "Bình", "Thực", "Tuấn", "Linh", "Thắng", "Nhật", "Mạnh" };
            string id = "";
            string name = "";
            string idDe = "";
            DateTime date = DateTime.Now.AddDays(this.GetRandom(1,300));
            string address = "";
            string[] add=new string[]{"Hà Nội","Hải Phòng","Ninh Bình","Thái Nguyên","Hồ Chí Minh"};
            string phone = "";
            string[] cns = new string[] { "PB001", "PB002", "PB003" };
            string queryFormat = "INSERT INTO NhanVien VALUES('{0}',N'{1}',N'{2}','{3}',N'{4}','{5}')";
            for (int i = 2; i < 60; i++)
            {
                id = "NV0" + i.ToString("x2");
                name = names[this.GetRandom(0, names.Length - 1)] + " " + name2[this.GetRandom(0, name2.Length - 1)] + " " + name3[this.GetRandom(0, name3.Length - 1)];
                idDe = cns[this.GetRandom(0, cns.Length - 1)];
                date = DateTime.Now.AddDays(this.GetRandom(1, 30)*(-1));
                address = add[this.GetRandom(0, add.Length - 1)];
                phone = "0121325454" + i.ToString("x2");
                string query = string.Format(queryFormat, id, name, idDe, date.ToString("yyyy-MM-dd HH:mm:ss"), address, phone);
                this.ExecuteNonQuery(query);
            }
        }

        public void ExecuteNonQuery(string query)
        {
            try
            {
                this.connection.Open();
                this.cmd = this.connection.CreateCommand();
                this.cmd.CommandText = query;
                this.cmd.CommandType = System.Data.CommandType.Text;
                this.cmd.ExecuteNonQuery();
                this.connection.Close();
            }
            catch (Exception e)
            {
                Common.log.Error("Get query: " + query, e);
            }
        }

        public DataTable ExecuteQuery(string query)
        {
            try
            {
                this.connection.Open();
                this.cmd = this.connection.CreateCommand();
                this.cmd.CommandText = query;
                this.cmd.CommandType = System.Data.CommandType.Text;
                SqlDataAdapter da = new SqlDataAdapter(this.cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                this.connection.Close();
                return dt;
            }
            catch (Exception e)
            {

                Common.log.Error("Get query: " + query, e);
                return new DataTable();
            }
            
        }

        public DataTable ExecuteStore(string nameStore)
        {
            try
            {
                this.connection.Open();
                this.cmd = this.connection.CreateCommand();
                this.cmd.CommandText = nameStore;
                this.cmd.CommandType = System.Data.CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(this.cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                this.connection.Close();
                return dt;
            }
            catch (Exception e)
            {
                Common.log.Error("Get from store: " + nameStore, e);
                return new DataTable();
            }
            
        }

        public DataTable ExecuteStore(string query,string[] names,object[] values)
        {
            try
            {
                if (names.Length != values.Length)
                {
                    throw new ArgumentOutOfRangeException("Số lượng tham số và giá trị không trùng khớp");
                }
                this.connection.Open();
                this.cmd = this.connection.CreateCommand();
                this.cmd.CommandText = query;
                this.cmd.CommandType = System.Data.CommandType.StoredProcedure;
                for (int i = 0; i < names.Length; i++)
                {
                    this.cmd.Parameters.Add(new SqlParameter(names[i], values[i]));
                }
                SqlDataAdapter da = new SqlDataAdapter(this.cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                this.connection.Close();
                return dt;
            }
            catch (Exception e)
            {
                Common.log.Error("Get store with param:" + query, e);
                return new DataTable();
            }
            
        }

        protected int GetRandom(int i,int j)
        {
            Random r = new Random();
            return r.Next(i, j);
        }

    }
}