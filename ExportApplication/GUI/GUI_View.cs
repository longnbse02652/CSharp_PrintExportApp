using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BLL;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExportApplication
{
    public partial class GUI_View : Form
    {
        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook = null;
        Excel.Worksheet xlWorkSheet = null;

        BLL_HandleFunc bll_handle = new BLL_HandleFunc();

        BLL_View bll_view = new BLL_View();
        DataTable dt = new DataTable();
        public static string xName = string.Empty;
        public GUI_View(string name)
        {
            InitializeComponent();
            groupBox2.Enabled = false;
            groupBox4.Enabled = false;
            groupBox1.Enabled = false;
            groupBox3.Enabled = false;
            tabPage1.Enabled = false;
            tabPage2.Enabled = false;
            tabPage3.Enabled = false;
            tabPage4.Enabled = false;

            xName = name;
            dt = bll_view.GetDataToView(name);
            //   MessageBox.Show(dt.Rows[0].Field<string>("BankName"));
            tb_IDCode.Text = dt.Rows[0].Field<string>("IDCode");
            tb_RomajiName.Text = dt.Rows[0].Field<string>("RomajiName");
            tb_FuriganaName.Text = dt.Rows[0].Field<string>("FuriganaName");
            cb_Sex.Text = dt.Rows[0].Field<string>("Sex");
            CheckDate(dtp_Birth, dt.Rows[0].Field<string>("Birth"));
            tb_Address1.Text = dt.Rows[0].Field<string>("Address1");
            cb_Address2.Text = dt.Rows[0].Field<string>("Address2");
            tb_Address3.Text = dt.Rows[0].Field<string>("Address3");
            cb_Address4.Text = dt.Rows[0].Field<string>("Address4");
            tb_ZipCode.Text = dt.Rows[0].Field<int?>("ZipCode").ToString();
            tb_MobliePhone.Text = dt.Rows[0].Field<string>("MobliePhone");
            tb_Phone.Text = dt.Rows[0].Field<string>("Phone");
            CheckDate(dtp_InCompanyDate,dt.Rows[0].Field<string>("InCompanyDate") );
            tb_CompanyCode.Text = dt.Rows[0].Field<string>("CompanyCode");
            tb_CompanyName.Text = dt.Rows[0].Field<string>("CompanyName");
            cb_WorkType.Text = dt.Rows[0].Field<string>("WorkType");
            cb_ClosingDate.Text = dt.Rows[0].Field<string>("ClosingDate");
            tb_Nationality.Text = dt.Rows[0].Field<string>("Nationality");
            cb_CardType.Text = dt.Rows[0].Field<string>("CardType");
            cb_OutTime.Text = dt.Rows[0].Field<string>("OutTime");
            CheckDate(dtp_CardTimeStart, dt.Rows[0].Field<string>("CardTime"));
            CheckDate(dtp_CardTimeOver, dt.Rows[0].Field<string>("CardTimeOut"));
            cb_Position.Text = dt.Rows[0].Field<string>("Position");
            tb_CreatePeople.Text = dt.Rows[0].Field<string>("CreatePeople");

            lb_chingin.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("chingin"));
            tb_HakenRyokin.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("HakenRyokin"));
            cb_HakenRyokinType.Text = dt.Rows[0].Field<string>("HakenRyokinType");
            cb_ShiharaiType.Text = dt.Rows[0].Field<string>("ShiharaiType");
            cb_Tax.Text = dt.Rows[0].Field<string>("Tax");
            cb_SalaryType.Text = dt.Rows[0].Field<string>("SalaryType");
            tb_BasicSalary.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("BasicSalary"));
            tb_SeikinTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("SeikinTeate"));
            tb_GaikinTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("GaikinTeate"));
            tb_GijutsuTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("GijutsuTeate"));
            tb_ShikakuTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("ShikakuTeate"));
            tb_YakushokuTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("YakushokuTeate"));
            tb_EigyoTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("EigyoTeate"));
            tb_KazokuTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("KazokuTeate"));
            tb_JutakuTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("JutakuTeate"));
            tb_BekkyoTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("BekkyoTeate"));
            tb_TsukinTeate.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("TsukinTeate"));
            tb_Park.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("Park"));
            tb_DormitoryFee.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("DormitoryFee"));
            tb_WaterFee.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("WaterFee"));
            
            cb_EmployStatus.Text = dt.Rows[0].Field<string>("EmployStatus");
            CheckDate(dtp_EmployTime1, dt.Rows[0].Field<string>("EmployTime1"));
            CheckDate(dtp_EmployTime2, dt.Rows[0].Field<string>("EmployTime2"));
            tb_BankName.Text = dt.Rows[0].Field<string>("BankName");
            cb_BankNameType.Text = dt.Rows[0].Field<string>("BankNameType");
            tb_BranchName.Text = dt.Rows[0].Field<string>("BranchName");
            cb_BranchNameType.Text = dt.Rows[0].Field<string>("BranchNameType");
            tb_AccountName.Text = dt.Rows[0].Field<string>("AccountName");
            tb_BranchCode.Text = dt.Rows[0].Field<string>("BranchCode");
            tb_BankCode.Text = dt.Rows[0].Field<string>("BankCode");
            tb_AccountCode1.Text = dt.Rows[0].Field<string>("AccountCode1");
            tb_AccountCode2.Text = dt.Rows[0].Field<string>("AccountCode2");
            tb_AccountCode3.Text = dt.Rows[0].Field<string>("AccountCode3");
            tb_AccountCode4.Text = dt.Rows[0].Field<string>("AccountCode4");
            tb_AccountCode5.Text = dt.Rows[0].Field<string>("AccountCode5");
            tb_AccountCode6.Text = dt.Rows[0].Field<string>("AccountCode6");
            tb_AccountCode7.Text = dt.Rows[0].Field<string>("AccountCode7");
            tb_AccountCode8.Text = dt.Rows[0].Field<string>("AccountCode8");
            cb_TravelType.Text = dt.Rows[0].Field<string>("TravelType");
            tb_HouseName.Text = dt.Rows[0].Field<string>("HouseName");
            tb_Room.Text = dt.Rows[0].Field<string>("Room");
            CheckDate(dtp_InHouseDate, dt.Rows[0].Field<string>("InHouseDate"));
            CheckDate(dtp_kouyouhoken, dt.Rows[0].Field<string>("Kouyouhoken"));
            CheckDate(dtp_shakaihoken, dt.Rows[0].Field<string>("Shakaihoken"));
            tb_DependentPeople.Text = dt.Rows[0].Field<int?>("DependentPeople").ToString();
            tb_ResidentPeople.Text = dt.Rows[0].Field<int?>("ResidentPeople").ToString();
            tb_HealthInsurancePeople.Text = dt.Rows[0].Field<int?>("HealthInsurancePeople").ToString();

            cb_ContractType.Text = dt.Rows[0].Field<string>("ContractType");
            cb_ContractRequire.Text = dt.Rows[0].Field<string>("ContractRequire");
            tb_MyCompany.Text = dt.Rows[0].Field<string>("MyCompany");
            tb_WorkContent.Text = dt.Rows[0].Field<string>("WorkContent");
            tb_WorkTime1.Text = dt.Rows[0].Field<string>("WorkTime1");
            tb_WorkTime2.Text = dt.Rows[0].Field<string>("WorkTime2");
            tb_WorkTime3.Text = dt.Rows[0].Field<string>("WorkTime3");
            tb_WorkTime4.Text = dt.Rows[0].Field<string>("WorkTime4");
            tb_RelaxTime.Text = dt.Rows[0].Field<string>("RelaxTime");
            cb_InsureCard.Text = dt.Rows[0].Field<string>("InsureCard");
            tb_PastCompany1.Text = dt.Rows[0].Field<string>("PastCompany1");
            cb_Nienhieu1.Text = dt.Rows[0].Field<string>("Nienhieu1");
            tb_BeginYear1.Text = dt.Rows[0].Field<int?>("BeginYear1").ToString();
            tb_BeginMonth1.Text = dt.Rows[0].Field<int?>("BeginMonth1").ToString();
            tb_EndYear1.Text = dt.Rows[0].Field<int?>("EndYear1").ToString();
            tb_EndMonth1.Text = dt.Rows[0].Field<int?>("EndMonth1").ToString();
            tb_PastCompany2.Text = dt.Rows[0].Field<string>("PastCompany2");
            cb_Nienhieu2.Text = dt.Rows[0].Field<string>("Nienhieu2");
            tb_BeginYear2.Text = dt.Rows[0].Field<int?>("BeginYear2").ToString();
            tb_BeginMonth2.Text = dt.Rows[0].Field<int?>("BeginMonth2").ToString();
            tb_EndYear2.Text = dt.Rows[0].Field<int?>("EndYear2").ToString();
            tb_EndMonth2.Text = dt.Rows[0].Field<int?>("EndMonth2").ToString();
            cb_PensionBook.Text = dt.Rows[0].Field<string>("PensionBook");
            tb_DependentPeopleKana1.Text = dt.Rows[0].Field<string>("DependentPeopleKana1");
            tb_DependentPeopleShimei1.Text = dt.Rows[0].Field<string>("DependentPeopleShimei1");
            CheckDate(dtp_DependentPeopleBirth1, dt.Rows[0].Field<string>("DependentPeopleBirth1"));
            tb_Relationship1.Text = dt.Rows[0].Field<string>("Relationship1");
            cb_Living1.Text = dt.Rows[0].Field<string>("Living1");
            tb_DependentPeopleKana2.Text = dt.Rows[0].Field<string>("DependentPeopleKana2");
            tb_DependentPeopleShimei2.Text = dt.Rows[0].Field<string>("DependentPeopleShimei2");
            CheckDate(dtp_DependentPeopleBirth2, dt.Rows[0].Field<string>("DependentPeopleBirth2"));
            tb_Relationship2.Text = dt.Rows[0].Field<string>("Relationship2");
            cb_Living2.Text = dt.Rows[0].Field<string>("Living2");
            tb_DependentPeopleKana3.Text = dt.Rows[0].Field<string>("DependentPeopleKana3");
            tb_DependentPeopleShimei3.Text = dt.Rows[0].Field<string>("DependentPeopleShimei3");
            CheckDate(dtp_DependentPeopleBirth3, dt.Rows[0].Field<string>("DependentPeopleBirth3"));
            tb_Relationship3.Text = dt.Rows[0].Field<string>("Relationship3");
            cb_Living3.Text = dt.Rows[0].Field<string>("Living3");
            tb_DependentPeopleKana4.Text = dt.Rows[0].Field<string>("DependentPeopleKana4");
            tb_DependentPeopleShimei4.Text = dt.Rows[0].Field<string>("DependentPeopleShimei4");
            CheckDate(dtp_DependentPeopleBirth4, dt.Rows[0].Field<string>("DependentPeopleBirth4"));
            tb_Relationship4.Text = dt.Rows[0].Field<string>("Relationship4");
            cb_Living4.Text = dt.Rows[0].Field<string>("Living4");
            tb_DependentPeopleKana5.Text = dt.Rows[0].Field<string>("DependentPeopleKana5");
            tb_DependentPeopleShimei5.Text = dt.Rows[0].Field<string>("DependentPeopleShimei5");
            CheckDate(dtp_DependentPeopleBirth5, dt.Rows[0].Field<string>("DependentPeopleBirth5"));
            tb_Relationship5.Text = dt.Rows[0].Field<string>("Relationship5");
            cb_Living5.Text = dt.Rows[0].Field<string>("Living5");
            tb_DependentPeopleKana6.Text = dt.Rows[0].Field<string>("DependentPeopleKana6");
            tb_DependentPeopleShimei6.Text = dt.Rows[0].Field<string>("DependentPeopleShimei6");
            CheckDate(dtp_DependentPeopleBirth6, dt.Rows[0].Field<string>("DependentPeopleBirth6"));
            tb_Relationship6.Text = dt.Rows[0].Field<string>("Relationship6");
            cb_Living6.Text = dt.Rows[0].Field<string>("Living6");

            tb_Trainsportation1.Text = dt.Rows[0].Field<string>("Trainsportation1");
            tb_BeginTrain1.Text = dt.Rows[0].Field<string>("BeginTrain1");
            tb_EndTrain1.Text = dt.Rows[0].Field<string>("EndTrain1");
            tb_MonthRegular1.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("MonthRegular1"));
            tb_Trainsportation2.Text = dt.Rows[0].Field<string>("Trainsportation2");
            tb_BeginTrain2.Text = dt.Rows[0].Field<string>("BeginTrain2");
            tb_EndTrain2.Text = dt.Rows[0].Field<string>("EndTrain2");
            tb_MonthRegular2.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("MonthRegular2"));
            tb_Trainsportation3.Text = dt.Rows[0].Field<string>("Trainsportation3");
            tb_BeginTrain3.Text = dt.Rows[0].Field<string>("BeginTrain3");
            tb_EndTrain3.Text = dt.Rows[0].Field<string>("EndTrain3");
            tb_MonthRegular3.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("MonthRegular3"));
            tb_Trainsportation4.Text = dt.Rows[0].Field<string>("Trainsportation4");
            tb_BeginTrain4.Text = dt.Rows[0].Field<string>("BeginTrain4");
            tb_EndTrain4.Text = dt.Rows[0].Field<string>("EndTrain4");
            tb_MonthRegular4.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("MonthRegular4"));
            cb_Carkm.Text = dt.Rows[0].Field<string>("Carkm");
            tb_CarMoney.Text = dt.Rows[0].Field<int?>("CarMoney").ToString();
            lb_TotalMoneyTrans.Text = string.Format("{0:n0}", dt.Rows[0].Field<int?>("TotalMoneyTrans"));

            int park, formitory, water;
            int.TryParse(dt.Rows[0].Field<int?>("Park").ToString(), out park);
            int.TryParse(dt.Rows[0].Field<int?>("DormitoryFee").ToString(), out formitory);
            int.TryParse(dt.Rows[0].Field<int?>("WaterFee").ToString(), out water);
            lb_tongtienkhautru.Text = string.Format("{0:n0}",park + formitory + water);
           
        }

        private void CheckDate(DateTimePicker dtp, string stringDate)
        {
            if (stringDate.Trim() != string.Empty)
            {
                dtp.Text = stringDate;
            }
        }

        private void bt_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dtp_Birth_ValueChanged(object sender, EventArgs e)
        {
            dtp_Birth.Format = DateTimePickerFormat.Long;
        }

        private void dtp_CardTimeStart_ValueChanged(object sender, EventArgs e)
        {
            dtp_CardTimeStart.Format = DateTimePickerFormat.Long;
        }

        private void dtp_CardTimeOver_ValueChanged(object sender, EventArgs e)
        {
            dtp_CardTimeOver.Format = DateTimePickerFormat.Long;
        }

        private void dtp_InCompanyDate_ValueChanged(object sender, EventArgs e)
        {
            dtp_InCompanyDate.Format = DateTimePickerFormat.Long;
        }

        private void dtp_EmployTime1_ValueChanged(object sender, EventArgs e)
        {
            dtp_EmployTime1.Format = DateTimePickerFormat.Long;
        }

        private void dtp_EmployTime2_ValueChanged(object sender, EventArgs e)
        {
            dtp_EmployTime2.Format = DateTimePickerFormat.Long;
        }

        private void dtp_kouyouhoken_ValueChanged(object sender, EventArgs e)
        {
            dtp_kouyouhoken.Format = DateTimePickerFormat.Long;
        }

        private void dtp_shakaihoken_ValueChanged(object sender, EventArgs e)
        {
            dtp_shakaihoken.Format = DateTimePickerFormat.Long;
        }

        private void dtp_InHouseDate_ValueChanged(object sender, EventArgs e)
        {
            dtp_InHouseDate.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth1_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth1.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth2_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth2.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth3_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth3.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth4_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth4.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth5_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth5.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth6_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth6.Format = DateTimePickerFormat.Long;
        }

        private void bt_Excel_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(new ThreadStart(SplashScreen));
            t.Start();
            Thread.Sleep(5000);

            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(path + @"\File\template_export.xls",
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //Export sheet nyushanaiyo
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                xlWorkSheet.Cells[8, "G"] = dt.Rows[0].Field<string>("IDCode");
                xlWorkSheet.Cells[10, "X"] = dt.Rows[0].Field<string>("RomajiName");
                xlWorkSheet.Cells[8, "X"] = dt.Rows[0].Field<string>("FuriganaName");
                xlWorkSheet.Cells[10, "AY"] = dt.Rows[0].Field<string>("Sex");
                if (dt.Rows[0].Field<string>("Birth") != string.Empty)
                {
                    xlWorkSheet.Cells[7, "DK"] = dt.Rows[0].Field<string>("Birth");
                }

                if (dt.Rows[0].Field<string>("InCompanyDate") != string.Empty)
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
                    if (temp_time1 != string.Empty)
                    {
                        string[] Time1_temps = temp_time1.Split('/');
                        xlWorkSheet.Cells[41, "AG"] = (Convert.ToInt32(Time1_temps[0]) - 1988).ToString();
                        xlWorkSheet.Cells[41, "AL"] = Time1_temps[1];
                        xlWorkSheet.Cells[41, "AP"] = Time1_temps[2];
                    }
                    string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                    if (temp_time2 != string.Empty)
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

                if (dt.Rows[0].Field<string>("Kouyouhoken") != string.Empty)
                {
                    xlWorkSheet.Cells[65, "AH"] = dt.Rows[0].Field<string>("Kouyouhoken");
                }
                if (dt.Rows[0].Field<string>("Shakaihoken") != string.Empty)
                {
                    xlWorkSheet.Cells[67, "AH"] = dt.Rows[0].Field<string>("Shakaihoken");
                }
                xlWorkSheet.Cells[67, "BC"] = dt.Rows[0].Field<int?>("DependentPeople");
                xlWorkSheet.Cells[67, "BM"] = dt.Rows[0].Field<int?>("ResidentPeople");
                xlWorkSheet.Cells[67, "BW"] = dt.Rows[0].Field<int?>("HealthInsurancePeople");

                xlWorkSheet.Cells[4, "BG"] = dt.Rows[0].Field<string>("CreatePeople");
                xlWorkSheet.Cells[3, "BG"] = dt.Rows[0].Field<string>("Position");


                ////////////////////////Export keiyaku
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
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
                if (joindate != " ")
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


                /////////////////////Export hoken
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
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



                ////Export koutsu
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);
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

                t.Abort(); //chỗ này để hủy luồng

                ////////// show promt to save file
                System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
                saveDlg.InitialDirectory = @"C:\";
                saveDlg.Filter = "Excel files (*.xls)|*.xls";
                saveDlg.FilterIndex = 0;
                saveDlg.RestoreDirectory = true;
                saveDlg.Title = "データ保存";
                saveDlg.FileName = dt.Rows[0].Field<string>("RomajiName") + "_" + DateTime.Now.ToString("yyyy-MM-dd") + "_";
                if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path1 = saveDlg.FileName;
                    xlWorkBook.SaveCopyAs(path1);
                    xlWorkBook.Saved = true;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.FinalReleaseComObject(xlWorkSheet);

                    xlWorkBook.Close(true, Type.Missing, Type.Missing);
                    Marshal.FinalReleaseComObject(xlWorkBook);

                    xlApp.Quit();
                    Marshal.FinalReleaseComObject(xlApp);
                    MessageBox.Show("出力完了");
                }
                else
                {
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                }
            }
            catch (Exception error)
            {
                // Cleanup Memory
                xlWorkBook.Close(0);
                xlApp.Quit();
                MessageBox.Show(error.Message, "エラー！出力できません！");
            }
        }

        private void SplashScreen() {
            Application.Run(new GUI_SplashScreen());
        }

        private void bt_Print_Click(object sender, EventArgs e)
        {
            GUI_Print gui_print = new GUI_Print(xName);
            gui_print.Show();
        }
    }
}
