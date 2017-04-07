using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DTO;
using System.Data.SqlClient;
using System.Data;
namespace DAL
{
   public class DAL_Dependent : DBConnection
    {
        public SqlCommand command;
        // load data len form dependent
        public DataTable DependentForm(string nameRomaji)
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter("select * from Information where RomajiName = '" + nameRomaji + "'", _cn);
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
        public bool Update(DTO_Dependent dto_dependent)
        {
            try
            {

                command = new SqlCommand("dbo.Dependent", _cn);
                command.CommandType = CommandType.StoredProcedure;

                command.Parameters.AddWithValue("@DependentPeopleKana1", dto_dependent.DependentPeopleKana1);
                command.Parameters.AddWithValue("@DependentPeopleKana2", dto_dependent.DependentPeopleKana2);
                command.Parameters.AddWithValue("@DependentPeopleKana3", dto_dependent.DependentPeopleKana3);
                command.Parameters.AddWithValue("@DependentPeopleKana4", dto_dependent.DependentPeopleKana4);
                command.Parameters.AddWithValue("@DependentPeopleKana5", dto_dependent.DependentPeopleKana5);
                command.Parameters.AddWithValue("@DependentPeopleKana6", dto_dependent.DependentPeopleKana6);
                command.Parameters.AddWithValue("@DependentPeopleShimei1", dto_dependent.DependentPeopleShimei1);
                command.Parameters.AddWithValue("@DependentPeopleShimei2", dto_dependent.DependentPeopleShimei2);
                command.Parameters.AddWithValue("@DependentPeopleShimei3", dto_dependent.DependentPeopleShimei3);
                command.Parameters.AddWithValue("@DependentPeopleShimei4", dto_dependent.DependentPeopleShimei4);
                command.Parameters.AddWithValue("@DependentPeopleShimei5", dto_dependent.DependentPeopleShimei5);
                command.Parameters.AddWithValue("@DependentPeopleShimei6", dto_dependent.DependentPeopleShimei6);
                command.Parameters.AddWithValue("@Relationship1", dto_dependent.Relationship1);
                command.Parameters.AddWithValue("@Relationship2", dto_dependent.Relationship2);
                command.Parameters.AddWithValue("@Relationship3", dto_dependent.Relationship3);
                command.Parameters.AddWithValue("@Relationship4", dto_dependent.Relationship4);
                command.Parameters.AddWithValue("@Relationship5", dto_dependent.Relationship5);
                command.Parameters.AddWithValue("@Relationship6", dto_dependent.Relationship6);
                command.Parameters.AddWithValue("@DependentPeopleBirth1", dto_dependent.DependentPeopleBirth1);
                command.Parameters.AddWithValue("@DependentPeopleBirth2", dto_dependent.DependentPeopleBirth2);
                command.Parameters.AddWithValue("@DependentPeopleBirth3", dto_dependent.DependentPeopleBirth3);
                command.Parameters.AddWithValue("@DependentPeopleBirth4", dto_dependent.DependentPeopleBirth4);
                command.Parameters.AddWithValue("@DependentPeopleBirth5", dto_dependent.DependentPeopleBirth5);
                command.Parameters.AddWithValue("@DependentPeopleBirth6", dto_dependent.DependentPeopleBirth6);
                command.Parameters.AddWithValue("@Living1", dto_dependent.Living1);
                command.Parameters.AddWithValue("@Living2", dto_dependent.Living2);
                command.Parameters.AddWithValue("@Living3", dto_dependent.Living3);
                command.Parameters.AddWithValue("@Living4", dto_dependent.Living4);
                command.Parameters.AddWithValue("@Living5", dto_dependent.Living5);
                command.Parameters.AddWithValue("@Living6", dto_dependent.Living6);
                command.Parameters.AddWithValue("@DependentPeople", dto_dependent.DependentPeople);
                command.Parameters.AddWithValue("@ResidentPeople", dto_dependent.ResidentPeople);
                command.Parameters.AddWithValue("@HealthInsurancePeople", dto_dependent.HealthInsurancePeople);
                command.Parameters.AddWithValue("@RomajiName", dto_dependent.RomajiName);
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
