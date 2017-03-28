using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class DAL_Edit : DBConnection
    {
        SqlDataReader myReader = null;
        public DataTable EditForm()
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter("select * from Edit", _cn);
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
