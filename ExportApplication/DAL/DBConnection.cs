using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class DBConnection
    {
        public SqlConnection _cn = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=dbExportExcelApp;Integrated Security=True");
    }
}
