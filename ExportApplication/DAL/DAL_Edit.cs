using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using System.Windows.Forms;


namespace DAL
{
    public class DAL_Edit : DBConnection
    {
        public SqlCommand command;
        // load data len form edit 
        public DataTable EditForm(string name)
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter("select * from Information where RomajiName = N'" + name + "'", _cn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
        }

        // Insert data to database
        public bool Update(DTO_Edit dto_Edit)
        {
            try
            {

                command = new SqlCommand("dbo.AddEdit", _cn);
                command.CommandType = CommandType.StoredProcedure;

                command.Parameters.AddWithValue("@IDCode", dto_Edit.IDCode);
                command.Parameters.AddWithValue("@RomajiName", dto_Edit.RomajiName);
                command.Parameters.AddWithValue("@FuriganaName", dto_Edit.FuriganaName);
                command.Parameters.AddWithValue("@Sex", dto_Edit.Sex);
                command.Parameters.AddWithValue("@CompanyName", dto_Edit.CompanyName);
                command.Parameters.AddWithValue("@CompanyCode", dto_Edit.CompanyCode);
                command.Parameters.AddWithValue("@ShiharaiType", dto_Edit.ShiharaiType);
                command.Parameters.AddWithValue("@Tax", dto_Edit.Tax);
                command.Parameters.AddWithValue("@TeateType", dto_Edit.TeateType);
                command.Parameters.AddWithValue("@Birth", dto_Edit.Birth);
                command.Parameters.AddWithValue("@Reason", dto_Edit.Reason);
                command.Parameters.AddWithValue("@ChangeDate", dto_Edit.ChangeDate);
                command.Parameters.AddWithValue("@ChangeDateFrom", dto_Edit.ChangeDateFrom);
                command.Parameters.AddWithValue("@ZipCode", dto_Edit.ZipCode);
                command.Parameters.AddWithValue("@Address1", dto_Edit.Address1);
                command.Parameters.AddWithValue("@Address2", dto_Edit.Address2);
                command.Parameters.AddWithValue("@Address3", dto_Edit.Address3);
                command.Parameters.AddWithValue("@Address4", dto_Edit.Address4);
                command.Parameters.AddWithValue("@Address5", dto_Edit.Address5);
                command.Parameters.AddWithValue("@TravelType", dto_Edit.TravelType);
                command.Parameters.AddWithValue("@EmployTime1", dto_Edit.EmployTime1);
                command.Parameters.AddWithValue("@EmployTime2", dto_Edit.EmployTime2);
                command.Parameters.AddWithValue("@CardType", dto_Edit.CardType);
                command.Parameters.AddWithValue("@CardTimeOut", dto_Edit.CardTimeOut);
                command.Parameters.AddWithValue("@CardTime", dto_Edit.CardTime);
                command.Parameters.AddWithValue("@WorkType", dto_Edit.WorkType);
                command.Parameters.AddWithValue("@ClosingDate", dto_Edit.ClosingDate);
                command.Parameters.AddWithValue("@HakenRyokin", dto_Edit.HakenRyokin);
                command.Parameters.AddWithValue("@HakenRyokinType", dto_Edit.HakenRyokinType);
                command.Parameters.AddWithValue("@Chingin", dto_Edit.Chingin);
                command.Parameters.AddWithValue("@ChinginType", dto_Edit.ChinginType);
                command.Parameters.AddWithValue("@TsukinTeate", dto_Edit.TsukinTeate);
                command.Parameters.AddWithValue("@Genkaritsu", dto_Edit.Genkaritsu);
                command.Parameters.AddWithValue("@TeateGaku", dto_Edit.TeateGaku);
                command.Parameters.AddWithValue("@KyuyoKojoGaku", dto_Edit.KyuyoKojoGaku);
                command.Parameters.AddWithValue("@WorkTime", dto_Edit.WorkTime);
                command.Parameters.AddWithValue("@BankName", dto_Edit.BankName);
                command.Parameters.AddWithValue("@BankNameType", dto_Edit.BankNameType);
                command.Parameters.AddWithValue("@BranchName", dto_Edit.BranchName);
                command.Parameters.AddWithValue("@BranchNameType", dto_Edit.BranchNameType);
                command.Parameters.AddWithValue("@AccountName", dto_Edit.AccountName);
                command.Parameters.AddWithValue("@BankCode", dto_Edit.BankCode);
                command.Parameters.AddWithValue("@BranchCode", dto_Edit.BranchCode);
                command.Parameters.AddWithValue("@AccountCode1", dto_Edit.AccountCode);
                command.Parameters.AddWithValue("@AccountCode2", dto_Edit.AccountCode1);
                command.Parameters.AddWithValue("@AccountCode3", dto_Edit.AccountCode2);
                command.Parameters.AddWithValue("@AccountCode4", dto_Edit.AccountCode3);
                command.Parameters.AddWithValue("@AccountCode5", dto_Edit.AccountCode4);
                command.Parameters.AddWithValue("@AccountCode6", dto_Edit.AccountCode5);
                command.Parameters.AddWithValue("@AccountCode7", dto_Edit.AccountCode6);
                command.Parameters.AddWithValue("@AccountCode8", dto_Edit.AccountCode7);
                command.Parameters.AddWithValue("@CompanyInsureDate", dto_Edit.CompanyInsureDate);
                command.Parameters.AddWithValue("@KoyoHokenDate", dto_Edit.KoyoHokenDate);
                command.Parameters.AddWithValue("@DependentPeople", dto_Edit.DependentPeople);
                command.Parameters.AddWithValue("@ResidentPeople", dto_Edit.ResidentPeople);
                command.Parameters.AddWithValue("@HealthInsurancePeople", dto_Edit.HealthInsurancePeople);
                _cn.Open();
                command.ExecuteNonQuery();
                _cn.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
                return false;
            }
        }


    }
}
