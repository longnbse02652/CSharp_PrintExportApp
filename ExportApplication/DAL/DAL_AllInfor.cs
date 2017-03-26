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
        DataTable dt = new DataTable();

        public DataTable GetDataToView()
        {
            try
            {
                adapter = new SqlDataAdapter("select RomajiName as '氏名', FuriganaName as 'ふりがな', Birth as '生年月日' from Information", _cn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
        }

        //Insert dữ liệu vô database
        public bool Insert(DTO_AllInfor dto_AllInfor) {
            try
            {
                command = new SqlCommand();
                command.CommandType = CommandType.Text;
                command.CommandText = "insert into Information values ('','"+dto_AllInfor.romaji+"','"+dto_AllInfor.furigana+"','','','','','','','','','','','','','','','','','','');"+
                                      "insert into Contract values ('"+dto_AllInfor.romaji+"','','','','','','','','');";
                command.Connection = _cn;

                _cn.Open();
                command.ExecuteNonQuery();
                _cn.Close();
                return true;
            }
            catch (Exception ex){
                MessageBox.Show(ex.Message, "Error Message");
                return false;
            }
        }


    }
}
