using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using System.Data.SqlClient;
using System.Data;

namespace DAL
{
    public class DAL_Infor : DBConnection
    {
        public DataTable GetToListView()
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter("select * from Information",_cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
        }


    }
}
