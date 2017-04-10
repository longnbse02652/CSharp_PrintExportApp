using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using BLL;
using System.IO;
using ExportApplication.Properties;
using System.Threading;

namespace ExportApplication
{
    public partial class GUI_Print : Form
    {
        BLL_Print bll_print = new BLL_Print();
        BLL_HandleFunc bll_handle = new BLL_HandleFunc();
        DataTable dt = new DataTable();

        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook = null;
        Excel.Worksheet xlWorkSheet = null;

        public GUI_Print(string getName)
        {
            InitializeComponent();
            dt = bll_print.GetDataToPrint(getName);
            //MessageBox.Show(dt.Rows[0].Field<string>("RomajiName"));
        }

        private object ValueOrDBNullIfZero(int val)
        {
            if (val == 0) return DBNull.Value;
            return val;
        }

        private void nyushanaiyousho()
        {

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[8, "G"] = dt.Rows[0].Field<string>("IDCode");
            xlWorkSheet.Cells[10, "X"] = dt.Rows[0].Field<string>("RomajiName");
            xlWorkSheet.Cells[8, "X"] = dt.Rows[0].Field<string>("FuriganaName");
            xlWorkSheet.Cells[10, "AY"] = dt.Rows[0].Field<string>("Sex");
            if (dt.Rows[0].Field<string>("Birth") != " ")
            {
                xlWorkSheet.Cells[7, "DK"] = dt.Rows[0].Field<string>("Birth");
            }

            if (dt.Rows[0].Field<string>("InCompanyDate") != " ")
            {
                xlWorkSheet.Cells[11, "DK"] = dt.Rows[0].Field<string>("InCompanyDate");
            }


            if (dt.Rows[0].Field<string>("Nationality") != "日本")
            {
                xlWorkSheet.Cells[11, "H"] = "日本以外は国名記入";
                xlWorkSheet.Cells[13, "H"] = dt.Rows[0].Field<string>("Nationality");
                string temp_zairyuukigen = dt.Rows[0].Field<string>("CardTimeOut");
                if (temp_zairyuukigen != " ")
                {
                    string[] temps = temp_zairyuukigen.Split('/');
                    xlWorkSheet.Cells[12, "BL"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[12, "BP"] = temps[1];
                    xlWorkSheet.Cells[12, "BT"] = temps[2];
                }
                switch (dt.Rows[0].Field<string>("CardType"))
                {
                    case "定住者":
                        xlWorkSheet.Cells[11, "AB"] = "☑ 定住者・永住者・特別永住・日本人配・永住配・技術";
                        break;
                    case "永住者":
                        xlWorkSheet.Cells[11, "AB"] = "定住者・☑ 永住者・特別永住・日本人配・永住配・技術";
                        break;
                    case "特別永住":
                        xlWorkSheet.Cells[11, "AB"] = "定住者・永住者・☑ 特別永住・日本人配・永住配・技術";
                        break;
                    case "日本人配":
                        xlWorkSheet.Cells[11, "AB"] = "定住者・永住者・特別永住・☑ 日本人配・永住配・技術";
                        break;
                    case "永住配":
                        xlWorkSheet.Cells[11, "AB"] = "定住者・永住者・特別永住・日本人配・☑ 永住配・技術";
                        break;
                    case "技術人文知識国際業務":
                        xlWorkSheet.Cells[11, "AB"] = "定住者・永住者・特別永住・日本人配・永住配・☑ 技術";
                        break;
                    case "留学":
                        xlWorkSheet.Cells[12, "AB"] = "人文知識国際業務・☑ 留学・就学・短期・家族・研修";
                        break;
                    case "就学":
                        xlWorkSheet.Cells[12, "AB"] = "人文知識国際業務・留学・☑ 就学・短期・家族・研修";
                        break;
                    case "短期":
                        xlWorkSheet.Cells[12, "AB"] = "人文知識国際業務・留学・就学・☑ 短期・家族・研修";
                        break;
                    case "家族":
                        xlWorkSheet.Cells[12, "AB"] = "人文知識国際業務・留学・就学・短期・☑ 家族・研修";
                        break;
                    case "研修":
                        xlWorkSheet.Cells[12, "AB"] = "人文知識国際業務・留学・就学・短期・家族・☑ 研修";
                        break;
                    default:
                        break;
                }
            }
            else
            {
                xlWorkSheet.Cells[11, "H"] = "日本";
            }

            xlWorkSheet.Cells[12, "BX"] = dt.Rows[0].Field<string>("OutTime");
            xlWorkSheet.Cells[17, "Y"] = dt.Rows[0].Field<string>("CompanyName");
            xlWorkSheet.Cells[19, "BI"] = dt.Rows[0].Field<string>("WorkType");
            xlWorkSheet.Cells[18, "BU"] = dt.Rows[0].Field<string>("ClosingDate");
            xlWorkSheet.Cells[24, "U"] = dt.Rows[0].Field<int?>("HakenRyokin");
            xlWorkSheet.Cells[24, "AH"] = dt.Rows[0].Field<string>("HakenRyokinType");
            xlWorkSheet.Cells[24, "AY"] = dt.Rows[0].Field<string>("ShiharaiType");
            xlWorkSheet.Cells[24, "BU"] = dt.Rows[0].Field<string>("Tax");
            xlWorkSheet.Cells[28, "AE"] = dt.Rows[0].Field<string>("SalaryType");
            xlWorkSheet.Cells[28, "AK"] = dt.Rows[0].Field<int?>("BasicSalary");
            xlWorkSheet.Cells[29, "AE"] = dt.Rows[0].Field<int?>("SeikinTeate");
            xlWorkSheet.Cells[30, "AE"] = dt.Rows[0].Field<int?>("GaikinTeate");
            xlWorkSheet.Cells[31, "AE"] = dt.Rows[0].Field<int?>("GijutsuTeate");
            xlWorkSheet.Cells[32, "AE"] = dt.Rows[0].Field<int?>("ShikakuTeate");
            xlWorkSheet.Cells[33, "AE"] = dt.Rows[0].Field<int?>("YakushokuTeate");
            xlWorkSheet.Cells[34, "AE"] = dt.Rows[0].Field<int?>("EigyoTeate");

            xlWorkSheet.Cells[35, "AE"] = dt.Rows[0].Field<int?>("KazokuTeate");
            xlWorkSheet.Cells[36, "AE"] = dt.Rows[0].Field<int?>("JutakuTeate");
            xlWorkSheet.Cells[37, "AE"] = dt.Rows[0].Field<int?>("BekkyoTeate");
            xlWorkSheet.Cells[38, "AM"] = dt.Rows[0].Field<int?>("TsukinTeate");

            xlWorkSheet.Cells[30, "BU"] = dt.Rows[0].Field<int?>("Park");
            xlWorkSheet.Cells[31, "BU"] = dt.Rows[0].Field<int?>("DormitoryFee");
            xlWorkSheet.Cells[32, "BU"] = dt.Rows[0].Field<int?>("WaterFee");

            xlWorkSheet.Cells[41, "G"] = dt.Rows[0].Field<string>("EmployStatus");
            if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
            {
                string temp_time1 = dt.Rows[0].Field<string>("EmployTime1");
                if (temp_time1 != " ")
                {
                    string[] Time1_temps = temp_time1.Split('/');
                    xlWorkSheet.Cells[41, "AG"] = (Convert.ToInt32(Time1_temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[41, "AL"] = Time1_temps[1];
                    xlWorkSheet.Cells[41, "AP"] = Time1_temps[2];
                }
                string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                if (temp_time2 != " ")
                {
                    string[] Time2_temps = temp_time2.Split('/');
                    xlWorkSheet.Cells[41, "AX"] = (Convert.ToInt32(Time2_temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[41, "BC"] = Time2_temps[1];
                    xlWorkSheet.Cells[41, "BG"] = Time2_temps[2];
                }

            }

            xlWorkSheet.Cells[53, "C"] = dt.Rows[0].Field<string>("BankName");
            xlWorkSheet.Cells[53, "AB"] = dt.Rows[0].Field<string>("BankNameType");
            xlWorkSheet.Cells[53, "AG"] = dt.Rows[0].Field<string>("BranchName");
            xlWorkSheet.Cells[53, "BB"] = dt.Rows[0].Field<string>("BranchNameType");
            xlWorkSheet.Cells[53, "BG"] = dt.Rows[0].Field<string>("AccountName");
            xlWorkSheet.Cells[56, "C"] = dt.Rows[0].Field<string>("BankCode");
            xlWorkSheet.Cells[56, "AG"] = dt.Rows[0].Field<string>("BranchCode");

            xlWorkSheet.Cells[56, "BI"] = dt.Rows[0].Field<string>("AccountCode1");
            xlWorkSheet.Cells[56, "BL"] = dt.Rows[0].Field<string>("AccountCode2");
            xlWorkSheet.Cells[56, "BO"] = dt.Rows[0].Field<string>("AccountCode3");
            xlWorkSheet.Cells[56, "BR"] = dt.Rows[0].Field<string>("AccountCode4");
            xlWorkSheet.Cells[56, "BU"] = dt.Rows[0].Field<string>("AccountCode5");
            xlWorkSheet.Cells[56, "BX"] = dt.Rows[0].Field<string>("AccountCode6");
            xlWorkSheet.Cells[56, "CA"] = dt.Rows[0].Field<string>("AccountCode7");
            xlWorkSheet.Cells[56, "CD"] = dt.Rows[0].Field<string>("AccountCode8");

            if (dt.Rows[0].Field<string>("TravelType") != "入寮")
            {
                xlWorkSheet.Cells[59, "F"] = "☑";
            }
            else
            {
                xlWorkSheet.Cells[61, "R"] = dt.Rows[0].Field<string>("HouseName");
                xlWorkSheet.Cells[61, "AY"] = dt.Rows[0].Field<string>("Room");
                if (dt.Rows[0].Field<string>("InHouseDate") != " ")
                {
                    string inhousedate = dt.Rows[0].Field<string>("InHouseDate");
                    string[] inhousedate_split = inhousedate.Split('/');
                    xlWorkSheet.Cells[61, "BR"] = (Convert.ToInt32(inhousedate_split[0]) - 1988).ToString();
                    xlWorkSheet.Cells[61, "BW"] = inhousedate_split[1];
                    xlWorkSheet.Cells[61, "CB"] = inhousedate_split[2];
                }
            }

            if (dt.Rows[0].Field<string>("Kouyouhoken") != " ")
            {
                xlWorkSheet.Cells[65, "AH"] = dt.Rows[0].Field<string>("Kouyouhoken");
            }
            if (dt.Rows[0].Field<string>("Shakaihoken") != " ")
            {
                xlWorkSheet.Cells[67, "AH"] = dt.Rows[0].Field<string>("Shakaihoken");
            }
            xlWorkSheet.Cells[67, "BC"] = dt.Rows[0].Field<int?>("DependentPeople");
            xlWorkSheet.Cells[67, "BM"] = dt.Rows[0].Field<int?>("ResidentPeople");
            xlWorkSheet.Cells[67, "BW"] = dt.Rows[0].Field<int?>("HealthInsurancePeople");

            xlWorkSheet.Cells[4, "BG"] = dt.Rows[0].Field<string>("CreatePeople");
            xlWorkSheet.Cells[3, "BG"] = dt.Rows[0].Field<string>("Position");

            // Print out 1 copy to the default printer:
            xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }

        private void keiyakusho()
        {
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            xlWorkSheet.Cells[5, "G"] = dt.Rows[0].Field<string>("RomajiName");
            xlWorkSheet.Cells[4, "G"] = dt.Rows[0].Field<string>("FuriganaName");
            xlWorkSheet.Cells[5, "T"] = dt.Rows[0].Field<string>("Sex");
            //Birthday
            string birth = dt.Rows[0].Field<string>("Birth");
            if (birth != " ")
            {
                string[] convert1 = birth.Split('/');
                string convert_birth = bll_handle.ConvertJapaneseCalendar(birth);
                xlWorkSheet.Cells[4, "Z"] = convert_birth.Substring(0, 2);
                xlWorkSheet.Cells[4, "AB"] = convert_birth.Substring(2, convert_birth.IndexOf("年") - 2);
                xlWorkSheet.Cells[4, "AD"] = convert1[1];
                xlWorkSheet.Cells[4, "AG"] = convert1[2];
            }

            //Zipcode
            string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
            if (zipcode.Length == 7)
            {
                string temp1_zipcode = zipcode.Substring(0, 3);
                string temp2_zipcode = zipcode.Substring(3, 4);
                xlWorkSheet.Cells[6, "H"] = temp1_zipcode;
                xlWorkSheet.Cells[6, "K"] = temp2_zipcode;
            }
            //Address
            xlWorkSheet.Cells[7, "G"] = dt.Rows[0].Field<string>("Address1");
            xlWorkSheet.Cells[7, "J"] = dt.Rows[0].Field<string>("Address2");
            xlWorkSheet.Cells[7, "K"] = dt.Rows[0].Field<string>("Address3");
            xlWorkSheet.Cells[7, "N"] = dt.Rows[0].Field<string>("Address4");
            xlWorkSheet.Cells[7, "O"] = dt.Rows[0].Field<string>("Address5");

            //Mobiphone
            string mobiphone = dt.Rows[0].Field<string>("MobliePhone");
            if (mobiphone.Length == 11)
            {
                string mobi1 = mobiphone.Substring(0, 3);
                string mobi2 = mobiphone.Substring(3, 4);
                string mobi3 = mobiphone.Substring(7, 4);
                xlWorkSheet.Cells[8, "W"] = mobi1;
                xlWorkSheet.Cells[8, "AA"] = mobi2;
                xlWorkSheet.Cells[8, "AD"] = mobi3;
            }
            //Phone
            string phone = dt.Rows[0].Field<string>("Phone");
            if (phone.Length == 10)
            {
                string phone1 = phone.Substring(0, 2);
                string phone2 = phone.Substring(2, 4);
                string phone3 = phone.Substring(6, 4);
                xlWorkSheet.Cells[8, "I"] = phone1;
                xlWorkSheet.Cells[8, "L"] = phone2;
                xlWorkSheet.Cells[8, "O"] = phone3;
            }
            //Join company Date
            string joindate = dt.Rows[0].Field<string>("InCompanyDate");
            if (dt.Rows[0].Field<string>("InCompanyDate") != " ")
            {
                string[] joindate_temps = joindate.Split('/');
                xlWorkSheet.Cells[10, "I"] = (Convert.ToInt32(joindate_temps[0]) - 1988).ToString();
                xlWorkSheet.Cells[10, "K"] = joindate_temps[1];
                xlWorkSheet.Cells[10, "M"] = joindate_temps[2];
            }
            //keiyaku time
            if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
            {
                xlWorkSheet.Cells[10, "Q"] = "□";
                xlWorkSheet.Cells[10, "V"] = "☑";
                string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                if (temp_time2 != " ")
                {
                    string[] Time2_temps = temp_time2.Split('/');
                    xlWorkSheet.Cells[10, "AB"] = (Convert.ToInt32(Time2_temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[10, "AD"] = Time2_temps[1];
                    xlWorkSheet.Cells[10, "AF"] = Time2_temps[2];
                }
            }
            //ContracType
            switch (dt.Rows[0].Field<string>("ContractType"))
            {
                case "自動的に更新する":
                    xlWorkSheet.Cells[11, "H"] = "1. 契約の更新の有無   ［☑ 自動的に更新する　・　更新する場合があり得る　・　契約の更新はしない］";
                    break;
                case "更新する場合があり得る":
                    xlWorkSheet.Cells[11, "H"] = "1. 契約の更新の有無   ［自動的に更新する　・　☑ 更新する場合があり得る　・　契約の更新はしない］";
                    break;
                case "契約の更新はしない":
                    xlWorkSheet.Cells[11, "H"] = "1. 契約の更新の有無   ［自動的に更新する　・　更新する場合があり得る　・　☑ 契約の更新はしない］";
                    break;
                default:
                    xlWorkSheet.Cells[11, "H"] = "1. 契約の更新の有無   ［自動的に更新する　・　更新する場合があり得る　・　契約の更新はしない］";
                    break;
            }
            //ContractRequire
            switch (dt.Rows[0].Field<string>("ContractRequire"))
            {
                case "契約期間満了時の業務量":
                    xlWorkSheet.Cells[12, "H"] = "2.契約の更新は次により判断する　　[☑ 契約期間満了時の業務量　　・勤務成績、態度 ・能力";
                    break;
                case "勤務成績、態度":
                    xlWorkSheet.Cells[12, "H"] = "2.契約の更新は次により判断する　　[契約期間満了時の業務量　・☑ 勤務成績、態度 ・能力";
                    break;
                case "能力":
                    xlWorkSheet.Cells[12, "H"] = "2.契約の更新は次により判断する　　[契約期間満了時の業務量　・勤務成績、態度 ・☑ 能力";
                    break;
                case "会社の経営状況":
                    xlWorkSheet.Cells[13, "H"] = "　・☑ 会社の経営状況 ・従事している業務の進捗状況　　・その他 (　　　　　　　　　　　　　　　　　　　　)　]";
                    break;
                case "従事している業務の進歩状況":
                    xlWorkSheet.Cells[13, "H"] = "　・会社の経営状況 ・☑ 従事している業務の進捗状況　　・☑ その他 (　　　　　　　　　　　　　　　　　　　　)　]";
                    break;
                default:
                    xlWorkSheet.Cells[13, "H"] = "　・会社の経営状況 ・従事している業務の進捗状況　　・その他 (　　　　　　　　　　　　　　　　　　　　)　]";
                    break;
            }

            //My company
            xlWorkSheet.Cells[14, "N"] = dt.Rows[0].Field<string>("MyCompany");
            xlWorkSheet.Cells[15, "G"] = dt.Rows[0].Field<string>("WorkContent");
            //Worktime
            xlWorkSheet.Cells[16, "L"] = dt.Rows[0].Field<string>("WorkTime1");
            xlWorkSheet.Cells[16, "N"] = dt.Rows[0].Field<string>("WorkTime2");
            xlWorkSheet.Cells[16, "U"] = dt.Rows[0].Field<string>("WorkTime3");
            xlWorkSheet.Cells[16, "W"] = dt.Rows[0].Field<string>("WorkTime4");
            xlWorkSheet.Cells[16, "AG"] = dt.Rows[0].Field<string>("RelaxTime");
            //賃金
            xlWorkSheet.Cells[24, "O"] = dt.Rows[0].Field<int?>("BasicSalary");
            xlWorkSheet.Cells[32, "P"] = dt.Rows[0].Field<int?>("SeikinTeate");
            xlWorkSheet.Cells[32, "AC"] = dt.Rows[0].Field<int?>("GaikinTeate");
            xlWorkSheet.Cells[33, "P"] = dt.Rows[0].Field<int?>("GijutsuTeate");
            xlWorkSheet.Cells[33, "AC"] = dt.Rows[0].Field<int?>("ShikakuTeate");
            xlWorkSheet.Cells[34, "P"] = dt.Rows[0].Field<int?>("YakushokuTeate");
            xlWorkSheet.Cells[34, "AC"] = dt.Rows[0].Field<int?>("EigyoTeate");
            xlWorkSheet.Cells[35, "P"] = dt.Rows[0].Field<int?>("KazokuTeate");
            xlWorkSheet.Cells[35, "AC"] = dt.Rows[0].Field<int?>("JutakuTeate");
            xlWorkSheet.Cells[36, "P"] = dt.Rows[0].Field<int?>("BekkyoTeate");
            xlWorkSheet.Cells[36, "AE"] = dt.Rows[0].Field<int?>("TsukinTeate");
            //寮費
            xlWorkSheet.Cells[40, "P"] = dt.Rows[0].Field<int?>("DormitoryFee");
            xlWorkSheet.Cells[42, "N"] = dt.Rows[0].Field<string>("ClosingDate");
            xlWorkSheet.Cells[40, "P"] = dt.Rows[0].Field<int?>("DormitoryFee");

            // Print out 1 copy to the default printer:
            xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }

        private void hoken()
        {

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);

            xlWorkSheet.Cells[2, "AP"] = dt.Rows[0].Field<string>("Position");
            xlWorkSheet.Cells[4, "AP"] = dt.Rows[0].Field<string>("CreatePeople");
            xlWorkSheet.Cells[11, "D"] = dt.Rows[0].Field<string>("IDCode");
            xlWorkSheet.Cells[12, "S"] = dt.Rows[0].Field<string>("RomajiName");
            xlWorkSheet.Cells[11, "S"] = dt.Rows[0].Field<string>("FuriganaName");
            xlWorkSheet.Cells[12, "AP"] = dt.Rows[0].Field<string>("Sex");
            xlWorkSheet.Cells[14, "S"] = dt.Rows[0].Field<string>("CompanyName");
            xlWorkSheet.Cells[15, "AL"] = dt.Rows[0].Field<string>("ClosingDate");
            //Birthday
            string birth = dt.Rows[0].Field<string>("Birth");
            if (birth != " ")
            {
                string[] convert1 = birth.Split('/');
                string convert_birth = bll_handle.ConvertJapaneseCalendar(birth);
                xlWorkSheet.Cells[12, "AS"] = convert_birth.Substring(0, 2);
                xlWorkSheet.Cells[12, "AT"] = convert_birth.Substring(2, convert_birth.IndexOf("年") - 2);
                xlWorkSheet.Cells[12, "AW"] = convert1[1];
                xlWorkSheet.Cells[12, "BA"] = convert1[2];

                //Age
                DateTime dt_birth = Convert.ToDateTime(birth);
                DateTime now = DateTime.Now;
                int age = now.Year - dt_birth.Year;
                if (now < dt_birth.AddYears(age)) age--;
                xlWorkSheet.Cells[12, "AL"] = age.ToString();
            }

            //Join company day
            string joindate = dt.Rows[0].Field<string>("InCompanyDate");
            if (joindate != " ")
            {
                string[] joindate_temps = joindate.Split('/');
                xlWorkSheet.Cells[15, "AT"] = (Convert.ToInt32(joindate_temps[0]) - 1988).ToString();
                xlWorkSheet.Cells[15, "AW"] = joindate_temps[1];
                xlWorkSheet.Cells[15, "BA"] = joindate_temps[2];
            }

            //koyouhoken
            string koyouhoken = dt.Rows[0].Field<string>("Kouyouhoken");
            if (koyouhoken != " ")
            {
                string[] koyouhoken_temp = koyouhoken.Split('/');
                xlWorkSheet.Cells[21, "P"] = (Convert.ToInt32(koyouhoken_temp[0]) - 1988).ToString();
                xlWorkSheet.Cells[21, "X"] = koyouhoken_temp[1];
                xlWorkSheet.Cells[21, "AF"] = koyouhoken_temp[2];
            }

            //ko co ng bao chung
            if (dt.Rows[0].Field<string>("InsureCard") != "有り" && dt.Rows[0].Field<string>("InsureCard") != string.Empty)
            {
                xlWorkSheet.Cells[24, "N"] = "□";
                xlWorkSheet.Cells[24, "X"] = "☑";
                xlWorkSheet.Cells[33, "D"] = dt.Rows[0].Field<string>("PastCompany1");
                xlWorkSheet.Cells[33, "AH"] = dt.Rows[0].Field<string>("Nienhieu1");
                xlWorkSheet.Cells[33, "AK"] = dt.Rows[0].Field<int?>("BeginYear1");
                xlWorkSheet.Cells[33, "AP"] = dt.Rows[0].Field<int?>("BeginMonth1");
                xlWorkSheet.Cells[33, "AV"] = dt.Rows[0].Field<int?>("EndYear1");
                xlWorkSheet.Cells[33, "BA"] = dt.Rows[0].Field<int?>("EndMonth1");

                xlWorkSheet.Cells[36, "D"] = dt.Rows[0].Field<string>("PastCompany2");
                xlWorkSheet.Cells[36, "AH"] = dt.Rows[0].Field<string>("Nienhieu2");
                xlWorkSheet.Cells[36, "AK"] = dt.Rows[0].Field<int?>("BeginYear2");
                xlWorkSheet.Cells[36, "AP"] = dt.Rows[0].Field<int?>("BeginMonth2");
                xlWorkSheet.Cells[36, "AV"] = dt.Rows[0].Field<int?>("EndYear2");
                xlWorkSheet.Cells[36, "BA"] = dt.Rows[0].Field<int?>("EndMonth2");
            }
            //Quoc tich va tu cach luu tru, thoi gian
            xlWorkSheet.Cells[39, "D"] = dt.Rows[0].Field<string>("Nationality");
            if (dt.Rows[0].Field<string>("CardTime") != " ")
            {
                xlWorkSheet.Cells[39, "AP"] = bll_handle.ConvertJapaneseCalendar(dt.Rows[0].Field<string>("CardTime"));
            }
            if (dt.Rows[0].Field<string>("CardTimeOut") != " ")
            {
                xlWorkSheet.Cells[39, "AW"] = bll_handle.ConvertJapaneseCalendar(dt.Rows[0].Field<string>("CardTimeOut"));
            }

            //shakaihoken
            string shakaihoken = dt.Rows[0].Field<string>("Shakaihoken");
            if (shakaihoken != " ")
            {
                string[] shakaihoken_temp = koyouhoken.Split('/');
                xlWorkSheet.Cells[45, "P"] = (Convert.ToInt32(shakaihoken_temp[0]) - 1988).ToString();
                xlWorkSheet.Cells[45, "X"] = shakaihoken_temp[1];
                xlWorkSheet.Cells[45, "AF"] = shakaihoken_temp[2];
            }
            //buu dien
            string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
            if (zipcode.Length == 7)
            {
                string temp1_zipcode = zipcode.Substring(0, 3);
                string temp2_zipcode = zipcode.Substring(3, 4);
                xlWorkSheet.Cells[49, "A"] = temp1_zipcode;
                xlWorkSheet.Cells[49, "G"] = temp2_zipcode;
            }
            //Address
            xlWorkSheet.Cells[49, "N"] = dt.Rows[0].Field<string>("Address1");
            xlWorkSheet.Cells[49, "S"] = dt.Rows[0].Field<string>("Address2");
            xlWorkSheet.Cells[49, "U"] = dt.Rows[0].Field<string>("Address3");
            xlWorkSheet.Cells[49, "AA"] = dt.Rows[0].Field<string>("Address4");
            xlWorkSheet.Cells[49, "AC"] = dt.Rows[0].Field<string>("Address5");
            //年金手帳
            if (dt.Rows[0].Field<string>("PensionBook") != "有り" && dt.Rows[0].Field<string>("PensionBook") != string.Empty)
            {
                xlWorkSheet.Cells[50, "N"] = "□";
                xlWorkSheet.Cells[50, "X"] = "☑";
            }
            //被扶養者
            xlWorkSheet.Cells[54, "K"] = dt.Rows[0].Field<string>("DependentPeopleKana1");
            xlWorkSheet.Cells[55, "K"] = dt.Rows[0].Field<string>("DependentPeopleShimei1");
            if (dt.Rows[0].Field<string>("DependentPeopleBirth1") != " ")
            {
                string depend1 = dt.Rows[0].Field<string>("DependentPeopleBirth1");
                string[] convert_depend1 = depend1.Split('/');
                string convert_japanStyle1 = bll_handle.ConvertJapaneseCalendar(depend1);
                xlWorkSheet.Cells[54, "AD"] = convert_japanStyle1.Substring(0, 2);　//平成、昭和
                xlWorkSheet.Cells[54, "AG"] = convert_japanStyle1.Substring(2, convert_japanStyle1.IndexOf("年") - 2); //年
                xlWorkSheet.Cells[54, "AK"] = convert_depend1[1]; //月
                xlWorkSheet.Cells[54, "AO"] = convert_depend1[2]; //日
                xlWorkSheet.Cells[54, "AS"] = dt.Rows[0].Field<string>("Relationship1");
                xlWorkSheet.Cells[54, "AX"] = dt.Rows[0].Field<string>("Living1");
            }


            xlWorkSheet.Cells[57, "K"] = dt.Rows[0].Field<string>("DependentPeopleKana2");
            xlWorkSheet.Cells[58, "K"] = dt.Rows[0].Field<string>("DependentPeopleShimei2");
            if (dt.Rows[0].Field<string>("DependentPeopleBirth2") != " ")
            {
                string depend2 = dt.Rows[0].Field<string>("DependentPeopleBirth2");
                string[] convert_depend2 = depend2.Split('/');
                string convert_japanStyle2 = bll_handle.ConvertJapaneseCalendar(depend2);
                xlWorkSheet.Cells[57, "AD"] = convert_japanStyle2.Substring(0, 2);　//平成、昭和
                xlWorkSheet.Cells[57, "AG"] = convert_japanStyle2.Substring(2, convert_japanStyle2.IndexOf("年") - 2); //年
                xlWorkSheet.Cells[57, "AK"] = convert_depend2[1]; //月
                xlWorkSheet.Cells[57, "AO"] = convert_depend2[2]; //日
                xlWorkSheet.Cells[57, "AS"] = dt.Rows[0].Field<string>("Relationship2");
                xlWorkSheet.Cells[57, "AX"] = dt.Rows[0].Field<string>("Living2");
            }

            xlWorkSheet.Cells[60, "K"] = dt.Rows[0].Field<string>("DependentPeopleKana3");
            xlWorkSheet.Cells[61, "K"] = dt.Rows[0].Field<string>("DependentPeopleShimei3");
            if (dt.Rows[0].Field<string>("DependentPeopleBirth3") != " ")
            {
                string depend3 = dt.Rows[0].Field<string>("DependentPeopleBirth3");
                string[] convert_depend3 = depend3.Split('/');
                string convert_japanStyle3 = bll_handle.ConvertJapaneseCalendar(depend3);
                xlWorkSheet.Cells[60, "AD"] = convert_japanStyle3.Substring(0, 2);　//平成、昭和
                xlWorkSheet.Cells[60, "AG"] = convert_japanStyle3.Substring(2, convert_japanStyle3.IndexOf("年") - 2); //年
                xlWorkSheet.Cells[60, "AK"] = convert_depend3[1]; //月
                xlWorkSheet.Cells[60, "AO"] = convert_depend3[2]; //日
                xlWorkSheet.Cells[60, "AS"] = dt.Rows[0].Field<string>("Relationship3");
                xlWorkSheet.Cells[60, "AX"] = dt.Rows[0].Field<string>("Living3");
            }

            xlWorkSheet.Cells[63, "K"] = dt.Rows[0].Field<string>("DependentPeopleKana4");
            xlWorkSheet.Cells[64, "K"] = dt.Rows[0].Field<string>("DependentPeopleShimei4");
            if (dt.Rows[0].Field<string>("DependentPeopleBirth4") != " ")
            {
                string depend4 = dt.Rows[0].Field<string>("DependentPeopleBirth4");
                string[] convert_depend4 = depend4.Split('/');
                string convert_japanStyle4 = bll_handle.ConvertJapaneseCalendar(depend4);
                xlWorkSheet.Cells[63, "AD"] = convert_japanStyle4.Substring(0, 2);　//平成、昭和
                xlWorkSheet.Cells[63, "AG"] = convert_japanStyle4.Substring(2, convert_japanStyle4.IndexOf("年") - 2); //年
                xlWorkSheet.Cells[63, "AK"] = convert_depend4[1]; //月
                xlWorkSheet.Cells[63, "AO"] = convert_depend4[2]; //日
                xlWorkSheet.Cells[63, "AS"] = dt.Rows[0].Field<string>("Relationship4");
                xlWorkSheet.Cells[63, "AX"] = dt.Rows[0].Field<string>("Living4");
            }

            xlWorkSheet.Cells[66, "K"] = dt.Rows[0].Field<string>("DependentPeopleKana5");
            xlWorkSheet.Cells[67, "K"] = dt.Rows[0].Field<string>("DependentPeopleShimei5");
            if (dt.Rows[0].Field<string>("DependentPeopleBirth5") != " ")
            {
                string depend5 = dt.Rows[0].Field<string>("DependentPeopleBirth5");
                string[] convert_depend5 = depend5.Split('/');
                string convert_japanStyle5 = bll_handle.ConvertJapaneseCalendar(depend5);
                xlWorkSheet.Cells[66, "AD"] = convert_japanStyle5.Substring(0, 2);　//平成、昭和
                xlWorkSheet.Cells[66, "AG"] = convert_japanStyle5.Substring(2, convert_japanStyle5.IndexOf("年") - 2); //年
                xlWorkSheet.Cells[66, "AK"] = convert_depend5[1]; //月
                xlWorkSheet.Cells[66, "AO"] = convert_depend5[2]; //日
                xlWorkSheet.Cells[66, "AS"] = dt.Rows[0].Field<string>("Relationship5");
                xlWorkSheet.Cells[66, "AX"] = dt.Rows[0].Field<string>("Living5");
            }

            xlWorkSheet.Cells[69, "K"] = dt.Rows[0].Field<string>("DependentPeopleKana6");
            xlWorkSheet.Cells[70, "K"] = dt.Rows[0].Field<string>("DependentPeopleShimei6");
            if (dt.Rows[0].Field<string>("DependentPeopleBirth6") != " ")
            {
                string depend6 = dt.Rows[0].Field<string>("DependentPeopleBirth6");
                string[] convert_depend6 = depend6.Split('/');
                string convert_japanStyle6 = bll_handle.ConvertJapaneseCalendar(depend6);
                xlWorkSheet.Cells[69, "AD"] = convert_japanStyle6.Substring(0, 2);　//平成、昭和
                xlWorkSheet.Cells[69, "AG"] = convert_japanStyle6.Substring(2, convert_japanStyle6.IndexOf("年") - 2); //年
                xlWorkSheet.Cells[69, "AK"] = convert_depend6[1]; //月
                xlWorkSheet.Cells[69, "AO"] = convert_depend6[2]; //日
                xlWorkSheet.Cells[69, "AS"] = dt.Rows[0].Field<string>("Relationship6");
                xlWorkSheet.Cells[69, "AX"] = dt.Rows[0].Field<string>("Living6");
            }

            // Print out 1 copy to the default printer:
            xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }

        public void koutsu()
        {

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);

            xlWorkSheet.Cells[3, "BC"] = dt.Rows[0].Field<string>("Position");
            xlWorkSheet.Cells[6, "BC"] = dt.Rows[0].Field<string>("CreatePeople");
            xlWorkSheet.Cells[21, "E"] = dt.Rows[0].Field<string>("IDCode");
            xlWorkSheet.Cells[21, "S"] = dt.Rows[0].Field<string>("RomajiName");
            xlWorkSheet.Cells[21, "AQ"] = dt.Rows[0].Field<string>("CompanyName");
            //通勤手当
            xlWorkSheet.Cells[87, "C"] = dt.Rows[0].Field<string>("Trainsportation1");
            xlWorkSheet.Cells[87, "P"] = dt.Rows[0].Field<string>("BeginTrain1");
            xlWorkSheet.Cells[87, "AD"] = dt.Rows[0].Field<string>("EndTrain1");
            xlWorkSheet.Cells[87, "AT"] = dt.Rows[0].Field<int?>("MonthRegular1");

            xlWorkSheet.Cells[90, "C"] = dt.Rows[0].Field<string>("Trainsportation2");
            xlWorkSheet.Cells[90, "P"] = dt.Rows[0].Field<string>("BeginTrain2");
            xlWorkSheet.Cells[90, "AD"] = dt.Rows[0].Field<string>("EndTrain2");
            xlWorkSheet.Cells[90, "AT"] = dt.Rows[0].Field<int?>("MonthRegular2");

            xlWorkSheet.Cells[93, "C"] = dt.Rows[0].Field<string>("Trainsportation3");
            xlWorkSheet.Cells[93, "P"] = dt.Rows[0].Field<string>("BeginTrain3");
            xlWorkSheet.Cells[93, "AD"] = dt.Rows[0].Field<string>("EndTrain3");
            xlWorkSheet.Cells[93, "AT"] = dt.Rows[0].Field<int?>("MonthRegular3");

            xlWorkSheet.Cells[96, "C"] = dt.Rows[0].Field<string>("Trainsportation4");
            xlWorkSheet.Cells[96, "P"] = dt.Rows[0].Field<string>("BeginTrain4");
            xlWorkSheet.Cells[96, "AD"] = dt.Rows[0].Field<string>("EndTrain4");
            xlWorkSheet.Cells[96, "AT"] = dt.Rows[0].Field<int?>("MonthRegular4");

            xlWorkSheet.Cells[99, "P"] = dt.Rows[0].Field<string>("Carkm");
            xlWorkSheet.Cells[99, "AT"] = dt.Rows[0].Field<int?>("CarMoney");
            xlWorkSheet.Cells[104, "AT"] = dt.Rows[0].Field<int?>("TotalMoneyTrans");


            // Print out 1 copy to the default printer:
            xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SplashScreen()
        {
            Application.Run(new GUI_SplashScreen());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(new ThreadStart(SplashScreen));
            t.Start();
            Thread.Sleep(5000);

            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path + @"\File\template.xls",
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            try
            {
                if (checkBox1.Checked == true)
                {
                    nyushanaiyousho();
                }
                if (checkBox2.Checked == true)
                {
                    keiyakusho();
                }
                if (checkBox3.Checked == true)
                {
                    hoken();
                }
                if (checkBox4.Checked == true)
                {
                    koutsu();
                }


                // Cleanup:
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);

                t.Abort(); //chỗ này để hủy luồng
                MessageBox.Show("印刷準備完了");

            }
            catch (Exception ex)
            {
                // Cleanup Memory
                xlWorkBook.Close(0);
                xlApp.Quit();
                MessageBox.Show(ex.Message, "エラー！印刷できません！");
            }

        }

    }
}
