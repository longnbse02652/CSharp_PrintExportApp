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
                command.Parameters.AddWithValue("@CardTimeStart", dto_AllInfor.cardTimeStart);
                command.Parameters.AddWithValue("@CardTimeOver", dto_AllInfor.cardTimeOver);
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
                
                command.Parameters.AddWithValue("@HakenRyokin", dto_AllInfor.hakenRyokin);
                command.Parameters.AddWithValue("@HakenRyokinType", dto_AllInfor.hakenRyokinType);
                command.Parameters.AddWithValue("@ShiharaiType", dto_AllInfor.shiharaiType);
                command.Parameters.AddWithValue("@Tax", dto_AllInfor.tax);
                command.Parameters.AddWithValue("@SalaryType", dto_AllInfor.salaryType);
                command.Parameters.AddWithValue("@BasicSalary", dto_AllInfor.basicSalary);
                command.Parameters.AddWithValue("@SeikinTeate", dto_AllInfor.seikinTeate);
                command.Parameters.AddWithValue("@GaikinTeate", dto_AllInfor.gaikinTeate);
                command.Parameters.AddWithValue("@GijutsuTeate", dto_AllInfor.gijutsuTeate);
                command.Parameters.AddWithValue("@ShikakuTeate", dto_AllInfor.shikakuTeate);
                command.Parameters.AddWithValue("@YakushokuTeate", dto_AllInfor.yakushokuTeate);
                command.Parameters.AddWithValue("@EigyoTeate", dto_AllInfor.eigyoTeate);
                command.Parameters.AddWithValue("@KazokuTeate", dto_AllInfor.kazokuTeate);
                command.Parameters.AddWithValue("@JutakuTeate", dto_AllInfor.jutakuTeate);
                command.Parameters.AddWithValue("@BekkyoTeate", dto_AllInfor.bekkyoTeate);
                command.Parameters.AddWithValue("@TsukinTeate", dto_AllInfor.tsukinTeate);
                command.Parameters.AddWithValue("@Park", dto_AllInfor.park);
                command.Parameters.AddWithValue("@DormitoryFee", dto_AllInfor.dormitoryFee);
                command.Parameters.AddWithValue("@WaterFee", dto_AllInfor.waterFee);
                command.Parameters.AddWithValue("@EmployStatus", dto_AllInfor.employStatus);
                command.Parameters.AddWithValue("@EmployTime1", dto_AllInfor.employTime1);
                command.Parameters.AddWithValue("@EmployTime2", dto_AllInfor.employTime2);
                command.Parameters.AddWithValue("@BankName", dto_AllInfor.bankName);
                command.Parameters.AddWithValue("@BankNameType", dto_AllInfor.bankNameType);
                command.Parameters.AddWithValue("@BranchName", dto_AllInfor.branchName);
                command.Parameters.AddWithValue("@BranchNameType", dto_AllInfor.branchNameType);
                command.Parameters.AddWithValue("@AccountName", dto_AllInfor.accountName);
                command.Parameters.AddWithValue("@BankCode", dto_AllInfor.bankCode);
                command.Parameters.AddWithValue("@BranchCode", dto_AllInfor.branchCode);
                command.Parameters.AddWithValue("@AccountCode1", dto_AllInfor.accountCode1);
                command.Parameters.AddWithValue("@AccountCode2", dto_AllInfor.accountCode2);
                command.Parameters.AddWithValue("@AccountCode3", dto_AllInfor.accountCode3);
                command.Parameters.AddWithValue("@AccountCode4", dto_AllInfor.accountCode4);
                command.Parameters.AddWithValue("@AccountCode5", dto_AllInfor.accountCode5);
                command.Parameters.AddWithValue("@AccountCode6", dto_AllInfor.accountCode6);
                command.Parameters.AddWithValue("@AccountCode7", dto_AllInfor.accountCode7);
                command.Parameters.AddWithValue("@AccountCode8", dto_AllInfor.accountCode8);
                command.Parameters.AddWithValue("@TravelType", dto_AllInfor.travelType);
                command.Parameters.AddWithValue("@HouseName", dto_AllInfor.houseName);
                command.Parameters.AddWithValue("@Room", dto_AllInfor.room);
                command.Parameters.AddWithValue("@InHouseDate", dto_AllInfor.inHouseDate);
                command.Parameters.AddWithValue("@Kouyouhoken", dto_AllInfor.kouyouhoken);
                command.Parameters.AddWithValue("@Shakaihoken", dto_AllInfor.shakaihoken);
                command.Parameters.AddWithValue("@DependentPeople", dto_AllInfor.dependentPeople);
                command.Parameters.AddWithValue("@ResidentPeople", dto_AllInfor.residentPeople);
                command.Parameters.AddWithValue("@HealthInsurancePeople", dto_AllInfor.healthInsurancePeople);
                command.Parameters.AddWithValue("@ContractType", dto_AllInfor.contractType);
                command.Parameters.AddWithValue("@ContractRequire", dto_AllInfor.contractRequire);
                command.Parameters.AddWithValue("@MyCompany", dto_AllInfor.myCompany);
                command.Parameters.AddWithValue("@WorkContent", dto_AllInfor.workContent);
                command.Parameters.AddWithValue("@WorkTime1", dto_AllInfor.workTime1);
                command.Parameters.AddWithValue("@WorkTime2", dto_AllInfor.workTime2);
                command.Parameters.AddWithValue("@WorkTime3", dto_AllInfor.workTime3);
                command.Parameters.AddWithValue("@WorkTime4", dto_AllInfor.workTime4);
                command.Parameters.AddWithValue("@RelaxTime", dto_AllInfor.relaxTime);

                command.Parameters.AddWithValue("@InsureCard", dto_AllInfor.insureCard);
                command.Parameters.AddWithValue("@PastCompany1", dto_AllInfor.pastCompany1);
                command.Parameters.AddWithValue("@Nienhieu1", dto_AllInfor.nienhieu1);
                command.Parameters.AddWithValue("@BeginYear1", dto_AllInfor.beginYear1);
                command.Parameters.AddWithValue("@BeginMonth1", dto_AllInfor.beginMonth1);
                command.Parameters.AddWithValue("@EndYear1", dto_AllInfor.endYear1);
                command.Parameters.AddWithValue("@EndMonth1", dto_AllInfor.endMonth1);
                command.Parameters.AddWithValue("@PastCompany2", dto_AllInfor.pastCompany2);
                command.Parameters.AddWithValue("@Nienhieu2", dto_AllInfor.nienhieu2);
                command.Parameters.AddWithValue("@BeginYear2", dto_AllInfor.beginYear2);
                command.Parameters.AddWithValue("@BeginMonth2", dto_AllInfor.beginMonth2);
                command.Parameters.AddWithValue("@EndYear2", dto_AllInfor.endYear2);
                command.Parameters.AddWithValue("@EndMonth2", dto_AllInfor.endMonth2);
                command.Parameters.AddWithValue("@PensionBook", dto_AllInfor.pensionBook);
                command.Parameters.AddWithValue("@DependentPeopleKana1", dto_AllInfor.dependentPeopleKana1);
                command.Parameters.AddWithValue("@DependentPeopleShimei1", dto_AllInfor.dependentPeopleShimei1);
                command.Parameters.AddWithValue("@DependentPeopleBirth1", dto_AllInfor.dependentPeopleBirth1);
                command.Parameters.AddWithValue("@Relationship1", dto_AllInfor.relationship1);
                command.Parameters.AddWithValue("@Living1", dto_AllInfor.living1);
                command.Parameters.AddWithValue("@DependentPeopleKana2", dto_AllInfor.dependentPeopleKana2);
                command.Parameters.AddWithValue("@DependentPeopleShimei2", dto_AllInfor.dependentPeopleShimei2);
                command.Parameters.AddWithValue("@DependentPeopleBirth2", dto_AllInfor.dependentPeopleBirth2);
                command.Parameters.AddWithValue("@Relationship2", dto_AllInfor.relationship1);
                command.Parameters.AddWithValue("@Living2", dto_AllInfor.living1);
                command.Parameters.AddWithValue("@DependentPeopleKana3", dto_AllInfor.dependentPeopleKana3);
                command.Parameters.AddWithValue("@DependentPeopleShimei3", dto_AllInfor.dependentPeopleShimei3);
                command.Parameters.AddWithValue("@DependentPeopleBirth3", dto_AllInfor.dependentPeopleBirth3);
                command.Parameters.AddWithValue("@Relationship3", dto_AllInfor.relationship1);
                command.Parameters.AddWithValue("@Living3", dto_AllInfor.living1);
                command.Parameters.AddWithValue("@DependentPeopleKana4", dto_AllInfor.dependentPeopleKana4);
                command.Parameters.AddWithValue("@DependentPeopleShimei4", dto_AllInfor.dependentPeopleShimei4);
                command.Parameters.AddWithValue("@DependentPeopleBirth4", dto_AllInfor.dependentPeopleBirth4);
                command.Parameters.AddWithValue("@Relationship4", dto_AllInfor.relationship1);
                command.Parameters.AddWithValue("@Living4", dto_AllInfor.living1);
                command.Parameters.AddWithValue("@DependentPeopleKana5", dto_AllInfor.dependentPeopleKana5);
                command.Parameters.AddWithValue("@DependentPeopleShimei5", dto_AllInfor.dependentPeopleShimei5);
                command.Parameters.AddWithValue("@DependentPeopleBirth5", dto_AllInfor.dependentPeopleBirth5);
                command.Parameters.AddWithValue("@Relationship5", dto_AllInfor.relationship1);
                command.Parameters.AddWithValue("@Living5", dto_AllInfor.living1);
                command.Parameters.AddWithValue("@DependentPeopleKana6", dto_AllInfor.dependentPeopleKana6);
                command.Parameters.AddWithValue("@DependentPeopleShimei6", dto_AllInfor.dependentPeopleShimei6);
                command.Parameters.AddWithValue("@DependentPeopleBirth6", dto_AllInfor.dependentPeopleBirth6);
                command.Parameters.AddWithValue("@Relationship6", dto_AllInfor.relationship1);
                command.Parameters.AddWithValue("@Living6", dto_AllInfor.living1);

                command.Parameters.AddWithValue("@Trainsportation1", dto_AllInfor.trainsportation1);
                command.Parameters.AddWithValue("@BeginTrain1", dto_AllInfor.beginTrain1);
                command.Parameters.AddWithValue("@EndTrain1", dto_AllInfor.endTrain1);
                command.Parameters.AddWithValue("@MonthRegular1", dto_AllInfor.monthRegular1);
                command.Parameters.AddWithValue("@Trainsportation2", dto_AllInfor.trainsportation2);
                command.Parameters.AddWithValue("@BeginTrain2", dto_AllInfor.beginTrain2);
                command.Parameters.AddWithValue("@EndTrain2", dto_AllInfor.endTrain2);
                command.Parameters.AddWithValue("@MonthRegular2", dto_AllInfor.monthRegular2);
                command.Parameters.AddWithValue("@Trainsportation3", dto_AllInfor.trainsportation3);
                command.Parameters.AddWithValue("@BeginTrain3", dto_AllInfor.beginTrain3);
                command.Parameters.AddWithValue("@EndTrain3", dto_AllInfor.endTrain3);
                command.Parameters.AddWithValue("@MonthRegular3", dto_AllInfor.monthRegular3);
                command.Parameters.AddWithValue("@Trainsportation4", dto_AllInfor.trainsportation4);
                command.Parameters.AddWithValue("@BeginTrain4", dto_AllInfor.beginTrain4);
                command.Parameters.AddWithValue("@EndTrain4", dto_AllInfor.endTrain4);
                command.Parameters.AddWithValue("@MonthRegular4", dto_AllInfor.monthRegular4);
                command.Parameters.AddWithValue("@Carkm", dto_AllInfor.carkm);
                command.Parameters.AddWithValue("@CarMoney", dto_AllInfor.carMoney);
                command.Parameters.AddWithValue("@TotalMoneyTrans", dto_AllInfor.totalMoneyTrans);

                command.Parameters.AddWithValue("@Reason", dto_AllInfor.reason);
                command.Parameters.AddWithValue("@ChangeDateFrom", dto_AllInfor.changeDateFrom);
                command.Parameters.AddWithValue("@ChangeDate", dto_AllInfor.changeDate);
                command.Parameters.AddWithValue("@Genkaritsu", dto_AllInfor.genkaritsu);
                command.Parameters.AddWithValue("@TeateGaku", dto_AllInfor.teateGaku);
                command.Parameters.AddWithValue("@AccountCode", dto_AllInfor.accountCode);
                command.Parameters.AddWithValue("@Chingin", dto_AllInfor.chingin);
                command.Parameters.AddWithValue("@ChinginType", dto_AllInfor.chinginType);
                command.Parameters.AddWithValue("@KyuyoKojoGaku", dto_AllInfor.kyuyoKojoGaku);
                command.Parameters.AddWithValue("@WorkTime", dto_AllInfor.workTime);
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
