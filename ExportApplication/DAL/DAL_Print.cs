using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class DAL_Print : DBConnection
    {
        public SqlDataAdapter adapter;
        public SqlCommand command;

        public DataTable GetDataToPrint(string name)
        {
            try
            {
                adapter = new SqlDataAdapter("select * from Information where RomajiName = N'" + name+ "'", _cn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
        }
    }
}
