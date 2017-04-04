using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DTO;


namespace DAL
{
   public class DAL_Travel:DBConnection
    {
        public SqlCommand command;
        // load data len form travel
        public DataTable TravelForm(string nameRomaji)
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter("select * from Information where RomajiName = N'" + nameRomaji + "'", _cn);
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
        public bool Update(DTO_Travel dto_travel)
        {
            try
            {

                command = new SqlCommand("dbo.AddTravel", _cn);
                command.CommandType = CommandType.StoredProcedure;

                command.Parameters.AddWithValue("@Trainsportation1", dto_travel.Trainsportation1);
                command.Parameters.AddWithValue("@BeginTrain1", dto_travel.BeginTrain1);
                command.Parameters.AddWithValue("@EndTrain1", dto_travel.EndTrain1);
                command.Parameters.AddWithValue("@MonthRegular1", dto_travel.MonthRegular1);
                command.Parameters.AddWithValue("@Trainsportation2", dto_travel.Trainsportation2);
                command.Parameters.AddWithValue("@BeginTrain2", dto_travel.BeginTrain2);
                command.Parameters.AddWithValue("@EndTrain2", dto_travel.EndTrain2);
                command.Parameters.AddWithValue("@MonthRegular2", dto_travel.MonthRegular2);

                command.Parameters.AddWithValue("@Trainsportation3", dto_travel.Trainsportation3);
                command.Parameters.AddWithValue("@BeginTrain3", dto_travel.BeginTrain3);
                command.Parameters.AddWithValue("@EndTrain3", dto_travel.EndTrain3);
                command.Parameters.AddWithValue("@MonthRegular3", dto_travel.MonthRegular3);
                command.Parameters.AddWithValue("@Trainsportation4", dto_travel.Trainsportation4);
                command.Parameters.AddWithValue("@BeginTrain4", dto_travel.BeginTrain4);
                command.Parameters.AddWithValue("@EndTrain4", dto_travel.EndTrain4);
                command.Parameters.AddWithValue("@MonthRegular4", dto_travel.MonthRegular4);

                command.Parameters.AddWithValue("@Carkm", dto_travel.Carkm);
                command.Parameters.AddWithValue("@CarMoney", dto_travel.CarMoney);
                command.Parameters.AddWithValue("@TotalMoneyTrans", dto_travel.TotalMoneyTrans);
                command.Parameters.AddWithValue("@RomajiName", dto_travel.RomajiName);
               
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
