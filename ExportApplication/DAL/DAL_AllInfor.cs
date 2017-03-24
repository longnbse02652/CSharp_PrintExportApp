using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace DAL
{
    public class DAL_AllInfor : DBConnection
    {
        public SqlDataAdapter adapter;
        public SqlCommand command;
        public DataTable GetToListView()
        {
            try
            {
                adapter = new SqlDataAdapter("select RomajiName, FuriganaName, Birth from Information", _cn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
        }

        DataTable dt = new DataTable();

        //Insert dữ liệu vô database
        public bool Insert(DTO_AllInfor dto_AllInfor) {
            try
            {
                adapter = new SqlDataAdapter("select * from Information", _cn); //con join tables nua, day chi test thoi
                adapter.Fill(dt);
                DataRow dr = dt.NewRow();
                dr["RomajiName"] = dto_AllInfor.romaji;
                dr["FuriganaName"] = dto_AllInfor.furigana;
                dr["Birth"] = dto_AllInfor.birth;

                dt.Rows.Add(dr);
                SqlCommandBuilder cm = new SqlCommandBuilder(adapter);
                adapter.Update(dt);
                return true;
            }
            catch (Exception ex){
                MessageBox.Show(ex.ToString());
                return false;
            }
        }


    }
}
