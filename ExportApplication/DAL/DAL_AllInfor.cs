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
                command = new SqlCommand("dbo.Addnew",_cn);
                command.CommandType = CommandType.StoredProcedure;

                command.Parameters.AddWithValue("@IDCode", dto_AllInfor.idCode);
                command.Parameters.AddWithValue("@RomajiName", dto_AllInfor.romaji);
                command.Parameters.AddWithValue("@FuriganaName", dto_AllInfor.furigana);
                command.Parameters.AddWithValue("@Sex", dto_AllInfor.sex);
                command.Parameters.AddWithValue("@Age", dto_AllInfor.age);
                command.Parameters.AddWithValue("@Birth", dto_AllInfor.birth);
                command.Parameters.AddWithValue("@Nationality", dto_AllInfor.nationality);
                command.Parameters.AddWithValue("@InCompanyDate", dto_AllInfor.inCompanyDate);
                command.Parameters.AddWithValue("@CardType", dto_AllInfor.cardType);
                command.Parameters.AddWithValue("@@CardTimeStart", dto_AllInfor.cardTime);
                command.Parameters.AddWithValue("@@CardTimeOver", dto_AllInfor.cardTimeOut);
                command.Parameters.AddWithValue("@OutTime",dto_AllInfor.outTime);
                command.Parameters.AddWithValue("@CompanyCode",dto_AllInfor.companyCode);
                command.Parameters.AddWithValue("@CompanyName",dto_AllInfor.companyName);
                command.Parameters.AddWithValue("@WorkType",dto_AllInfor.workType);
                command.Parameters.AddWithValue("@ClosingDate",dto_AllInfor.closingDate);
                command.Parameters.AddWithValue("@ZipCode",dto_AllInfor.zipCode);
                command.Parameters.AddWithValue("@Address",dto_AllInfor.address);
                command.Parameters.AddWithValue("@MobliePhone",dto_AllInfor.mobliePhone);
                command.Parameters.AddWithValue("@Phone",dto_AllInfor.phone);
                command.Parameters.AddWithValue("@CreatePeople",dto_AllInfor.createPeople);
                command.Parameters.AddWithValue("@Position",dto_AllInfor.position);
                //
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
