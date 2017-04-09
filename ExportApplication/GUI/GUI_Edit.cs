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
using DTO;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportApplication
{
    public partial class GUI_Edit : Form
    {
        BLL_Edit bll_edit = new BLL_Edit();
        BLL_HandleFunc bll_handleFunc = new BLL_HandleFunc();

        //////////////////////////////////////////ENABLE AND DISABLE IN GUI////////////////////////////////
        public void TakeThis(IList<string> str)
        {
            TB_Reason.Enabled = true;
            DTP_ChangeDate.Enabled = true;
            DTP_ChangeDateFrom.Enabled = true;
            for (int i = 0; i < str.Count; i++)
            {
                switch (str[i])
                {
                    case "氏名":
                        TB_RomajiName.Enabled = true;
                        TB_FuriganaName.Enabled = true;
                        label2.Font = new Font(label2.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label2.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "住民票住所":
                        TB_ZipCode.Enabled = true;
                        TB_Address1.Enabled = true;
                        TB_Address3.Enabled = true;
                        TB_Address5.Enabled = true;
                        CB_Address2.Enabled = true;
                        CB_Address4.Enabled = true;
                        label24.Font = new Font(label24.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label24.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "居所住所":
                        break;
                    case "通勤/入寮":
                        CB_TravelType.Enabled = true;
                        label22.Font = new Font(label22.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label22.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "雇用期間":
                        DTP_EmployTime1.Enabled = true;
                        DTP_EmployTime2.Enabled = true;
                        label21.Font = new Font(label21.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label21.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "在留資格":
                        CB_CardType.Enabled = true;
                        label20.Font = new Font(label20.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label20.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "在留期間":
                        DTP_CardTimeStart.Enabled = true;
                        DTP_CardTimeOver.Enabled = true;
                        label14.Font = new Font(label14.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label14.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "企業名":
                        TB_CompanyName.Enabled = true;
                        TB_CompanyCode.Enabled = true;
                        label5.Font = new Font(label5.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label5.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "締日":
                        CB_ClosingDate.Enabled = true;
                        label8.Font = new Font(label8.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label8.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "就労形態":
                        CB_WorkType.Enabled = true;
                        label15.Font = new Font(label15.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label15.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "賃金支払形態/税適":
                        CB_ShiharaiType.Enabled = true;
                        label7.Font = new Font(label7.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label7.ForeColor = System.Drawing.Color.Red;
                        CB_Tax.Enabled = true;
                        label26.Font = new Font(label26.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label26.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "派遣・請金":
                        TB_HakenRyokin.Enabled = true;
                        CB_HakenRyokinType.Enabled = true;
                        label27.Font = new Font(label27.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label27.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "賃金":
                        TB_Chingin.Enabled = true;
                        CB_ChinginType.Enabled = true;
                        label30.Font = new Font(label30.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label30.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "通勤手当":
                        TB_TsukinTeate.Enabled = true;
                        label32.Font = new Font(label32.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label32.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "手当額":
                        TB_TeateGaku.Enabled = true;
                        CB_TeateType.Enabled = true;
                        label34.Font = new Font(label34.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label34.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "給与控除額":
                        TB_KyuyoKojoGaku.Enabled = true;
                        label36.Font = new Font(label36.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label36.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "労働時間":
                        TB_WorkTime.Enabled = true;
                        label38.Font = new Font(label38.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label38.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "銀行/支店名":
                        TB_BankCode.Enabled = true;
                        label39.Font = new Font(label39.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label39.ForeColor = System.Drawing.Color.Red;
                        TB_BranchCode.Enabled = true;
                        label40.Font = new Font(label40.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label40.ForeColor = System.Drawing.Color.Red;
                        TB_BankName.Enabled = true;
                        CB_BankNameType.Enabled = true;
                        label41.Font = new Font(label41.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label41.ForeColor = System.Drawing.Color.Red;
                        TB_BranchName.Enabled = true;
                        CB_BranchNameType.Enabled = true;
                        label42.Font = new Font(label42.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label42.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "口座名義（カナ）":
                        TB_AccountName.Enabled = true;
                        label43.Font = new Font(label43.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label43.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "口座番号":
                        TB_AccountCode.Enabled = true;
                        TB_AccountCode1.Enabled = true;
                        TB_AccountCode2.Enabled = true;
                        TB_AccountCode3.Enabled = true;
                        TB_AccountCode4.Enabled = true;
                        TB_AccountCode5.Enabled = true;
                        TB_AccountCode6.Enabled = true;
                        TB_AccountCode7.Enabled = true;
                        label44.Font = new Font(label44.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label44.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "保険":
                        DTP_KoyoHokenDate.Enabled = true;
                        label45.Font = new Font(label45.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label45.ForeColor = System.Drawing.Color.Red;
                        DTP_CompanyInsureDate.Enabled = true;
                        label46.Font = new Font(label46.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label46.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "扶養人数":
                        TB_DependentPeople.Enabled = true;
                        TB_ResidentPeople.Enabled = true;
                        TB_HealthInsurancePeople.Enabled = true;
                        label47.Font = new Font(label47.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label47.ForeColor = System.Drawing.Color.Red;
                        label48.Font = new Font(label48.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label48.ForeColor = System.Drawing.Color.Red;
                        label49.Font = new Font(label49.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label49.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "性別":
                        CB_Sex.Enabled = true;
                        label6.Font = new Font(label6.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label6.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "生年月日":
                        DTP_Birth.Enabled = true;
                        label6.Font = new Font(label6.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label6.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "社員ＣＤ":
                        TB_IDCode.Enabled = true;
                        label1.Font = new Font(label1.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label1.ForeColor = System.Drawing.Color.Red;
                        break;
                    case "企業ＣＤ":
                        TB_CompanyCode.Enabled = true;
                        label4.Font = new Font(label4.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label4.ForeColor = System.Drawing.Color.Red;
                        break;
                }
            }
        }
        public GUI_Edit()
        {
            InitializeComponent();
           
           // CB_Sex.Text = dt.Rows[0].Field<string>("Sex");
        }
        // get name lay tu dataGridView de gui len DAL
        public static string name;
        public void funData(string text)
        {
            name = text;
            
        }

        ///////////////////////////////////////LOAD DATA FROM DATABASE TO GUI///////////////////////////////////

        public delegate void delPassData(string text);
        string old_tsukinTeate = string.Empty;
        string old_DependentPeople = string.Empty;
        string old_ResidentPeople = string.Empty;
        string old_HealthInsurancePeople = string.Empty;
        int old_SeikinTeate, old_GaikinTeate, old_GijutsuTeate, old_ShikakuTeate, old_YakushokuTeate, old_EigyoTeate, old_JutakuTeate, old_BekkyoTeate, old_KazokuTeate;

       
        private void GUI_Edit_Load(object sender, EventArgs e)
        {
            DataTable dt = bll_edit.EditForm(name);
            old_tsukinTeate = dt.Rows[0].Field<int?>("TotalMoneyTrans").ToString();
            old_DependentPeople = dt.Rows[0].Field<int?>("DependentPeople").ToString();
            old_ResidentPeople = dt.Rows[0].Field<int?>("ResidentPeople").ToString();
            old_HealthInsurancePeople = dt.Rows[0].Field<int?>("HealthInsurancePeople").ToString();
            foreach (DataRow row in dt.Rows)
            {
                // Luu gia tro old de luc imsert data lay giu lieu
                object SeikinTeate = row["SeikinTeate"];
                object GaikinTeate = row["GaikinTeate"];
                object GijutsuTeate = row["GijutsuTeate"];
                object ShikakuTeate = row["ShikakuTeate"];
                object YakushokuTeate = row["YakushokuTeate"];
                object EigyoTeate = row["EigyoTeate"];
                object JutakuTeate = row["JutakuTeate"];
                object BekkyoTeate = row["BekkyoTeate"];
                object KazokuTeate = row["KazokuTeate"];
                if (SeikinTeate == DBNull.Value)
                {
                    old_SeikinTeate = 0;
                }
                else
                {
                    old_SeikinTeate = dt.Rows[0].Field<int>("SeikinTeate");
                }

                if (GaikinTeate == DBNull.Value)
                {
                    old_GaikinTeate = 0;
                }
                else
                {
                    old_GaikinTeate = dt.Rows[0].Field<int>("GaikinTeate");
                }

                if (GijutsuTeate == DBNull.Value)
                {
                    old_GijutsuTeate = 0;
                }
                else
                {
                    old_GijutsuTeate = dt.Rows[0].Field<int>("GijutsuTeate");
                }
                if (ShikakuTeate == DBNull.Value)
                {
                    old_ShikakuTeate = 0;
                }
                else
                {
                    old_ShikakuTeate = dt.Rows[0].Field<int>("ShikakuTeate");
                }
                if (YakushokuTeate == DBNull.Value)
                {
                    old_YakushokuTeate = 0;
                }
                else
                {
                    old_YakushokuTeate = dt.Rows[0].Field<int>("YakushokuTeate");
                }

                if (EigyoTeate == DBNull.Value)
                {
                    old_EigyoTeate = 0;
                }
                else
                {
                    old_EigyoTeate = dt.Rows[0].Field<int>("EigyoTeate");
                }
                if (JutakuTeate == DBNull.Value)
                {
                    old_JutakuTeate = 0;
                }
                else
                {
                    old_JutakuTeate = dt.Rows[0].Field<int>("JutakuTeate");
                }
                if (BekkyoTeate == DBNull.Value)
                {
                    old_BekkyoTeate = 0;
                }
                else
                {
                    old_BekkyoTeate = dt.Rows[0].Field<int>("BekkyoTeate");
                }

                if (KazokuTeate == DBNull.Value)
                {
                    old_KazokuTeate = 0;
                }
                else
                {
                    KazokuTeate = dt.Rows[0].Field<int>("KazokuTeate");
                }
                TB_RomajiName.Text = dt.Rows[0].Field<string>("RomajiName");
                TB_IDCode.Text = dt.Rows[0].Field<string>("IDCode");
                TB_FuriganaName.Text = dt.Rows[0].Field<string>("FuriganaName");
                TB_CompanyCode.Text = dt.Rows[0].Field<string>("CompanyCode");
                TB_CompanyName.Text = dt.Rows[0].Field<string>("CompanyName");
                TB_Reason.Text = dt.Rows[0].Field<string>("Reason");
                TB_Address1.Text = dt.Rows[0].Field<string>("Address1");// la cho hang t2 trong bang edit
                TB_Address3.Text = dt.Rows[0].Field<string>("Address3");
                //   TB_Address1.Text = dt.Rows[0].Field<string>("Address1");
                TB_Address5.Text = dt.Rows[0].Field<string>("Address5");
                TB_TeateGaku.Text = dt.Rows[0].Field<string>("TeateGaku");
                TB_BankName.Text = dt.Rows[0].Field<string>("BankName");
                TB_BranchName.Text = dt.Rows[0].Field<string>("BranchName");
                TB_AccountName.Text = dt.Rows[0].Field<string>("AccountName");
                TB_BankCode.Text = dt.Rows[0].Field<string>("BankCode");
                TB_BranchCode.Text = dt.Rows[0].Field<string>("BranchCode");
                TB_AccountCode.Text = dt.Rows[0].Field<string>("AccountCode1");
                TB_AccountCode1.Text = dt.Rows[0].Field<string>("AccountCode2");
                TB_AccountCode2.Text = dt.Rows[0].Field<string>("AccountCode3");
                TB_AccountCode3.Text = dt.Rows[0].Field<string>("AccountCode4");
                TB_AccountCode4.Text = dt.Rows[0].Field<string>("AccountCode5");
                TB_AccountCode5.Text = dt.Rows[0].Field<string>("AccountCode6");
                TB_AccountCode6.Text = dt.Rows[0].Field<string>("AccountCode7");
                TB_AccountCode7.Text = dt.Rows[0].Field<string>("AccountCode8");
                TB_KyuyoKojoGaku.Text = dt.Rows[0].Field<string>("DormitoryFee");
                CB_TravelType.Text = dt.Rows[0].Field<string>("TravelType");
                CB_BranchNameType.Text = dt.Rows[0].Field<string>("BranchNameType");
                CB_BankNameType.Text = dt.Rows[0].Field<string>("BankNameType");
                CB_TeateType.Text = dt.Rows[0].Field<string>("TeateType");
                CB_ChinginType.Text = dt.Rows[0].Field<string>("ChinginType");
                CB_HakenRyokinType.Text = dt.Rows[0].Field<string>("HakenRyokinType");
                CB_ClosingDate.Text = dt.Rows[0].Field<string>("ClosingDate");
                CB_WorkType.Text = dt.Rows[0].Field<string>("WorkType");
                CB_CardType.Text = dt.Rows[0].Field<string>("CardType");
                CB_ShiharaiType.Text = dt.Rows[0].Field<string>("ShiharaiType");
                CB_Sex.Text = dt.Rows[0].Field<string>("Sex");
                CB_Tax.Text = dt.Rows[0].Field<string>("Tax");
                CB_Address2.Text = dt.Rows[0].Field<string>("Address2");
                CB_Address4.Text = dt.Rows[0].Field<string>("Address4");

                object changedate = row["ChangeDate"];
                object changedatefrom = row["ChangeDateFrom"];
                object birth = row["Birth"];
                object employtime1 = row["employtime1"];
                object employtime2 = row["employtime2"];
                object cardtimeover = row["CardTimeOut"];
                object cardtimestart = row["CardTime"];
                object hakenryokin = row["HakenRyokin"];
                object chingin = row["Chingin"];
                object tsukinteate = row["TotalMoneyTrans"];
                object kyuyokojogaku = row["KyuyoKojoGaku"];
                object worktime = row["WorkTime"];
                object companyinsuredate = row["Shakaihoken"];
                object koyohokendate = row["Kouyouhoken"];
                object dependentpeople = row["DependentPeople"];
                object residentpeople = row["ResidentPeople"];
                object healthinsurancepeople = row["HealthInsurancePeople"];
                object zipcode = row["ZipCode"];
                object zipcode1 = row["ZipCode"];

                if (string.IsNullOrEmpty(row["ChangeDate"].ToString()) || row["ChangeDate"].ToString() == " ")
                {
                    DTP_ChangeDate.Format = DateTimePickerFormat.Custom;
                    DTP_ChangeDate.CustomFormat = " ";
                }
                else { DTP_ChangeDate.Text = dt.Rows[0].Field<string>("ChangeDate"); }

                if (changedatefrom == DBNull.Value || row["ChangeDateFrom"].ToString() == " ")
                {
                    DTP_ChangeDateFrom.Format = DateTimePickerFormat.Custom;
                    DTP_ChangeDateFrom.CustomFormat = " ";
                }
                else { DTP_ChangeDateFrom.Text = dt.Rows[0].Field<string>("ChangeDateFrom"); }

                if (birth == DBNull.Value || row["Birth"].ToString() == " ")
                {
                    DTP_Birth.Format = DateTimePickerFormat.Custom;
                    DTP_Birth.CustomFormat = " ";
                }
                else { DTP_Birth.Text = dt.Rows[0].Field<string>("Birth"); }

                if (string.IsNullOrEmpty(row["EmployTime1"].ToString()) || row["EmployTime1"].ToString() == " ")
                {
                    DTP_EmployTime1.Format = DateTimePickerFormat.Custom;
                    DTP_EmployTime1.CustomFormat = " ";
                }
                else { DTP_EmployTime1.Text = dt.Rows[0].Field<string>("EmployTime1"); }

                if (employtime2 == DBNull.Value || row["EmployTime2"].ToString() == " ")
                {
                    DTP_EmployTime2.Format = DateTimePickerFormat.Custom;
                    DTP_EmployTime2.CustomFormat = " ";
                }
                else { DTP_EmployTime2.Text = dt.Rows[0].Field<string>("EmployTime2"); }

                if (cardtimeover == DBNull.Value || row["CardTimeOut"].ToString() == " ")
                {
                    DTP_CardTimeOver.Format = DateTimePickerFormat.Custom;
                    DTP_CardTimeOver.CustomFormat = " ";
                }
                else { DTP_CardTimeOver.Text = dt.Rows[0].Field<string>("CardTimeOut"); }

                if (cardtimestart == DBNull.Value || row["CardTime"].ToString() == " ")
                {
                    DTP_CardTimeStart.Format = DateTimePickerFormat.Custom;
                    DTP_CardTimeStart.CustomFormat = " ";
                }
                else { DTP_CardTimeStart.Text = dt.Rows[0].Field<string>("CardTime"); }

                if (hakenryokin == DBNull.Value)
                {
                    TB_HakenRyokin.Text = " ";
                }
                else { TB_HakenRyokin.Text = dt.Rows[0].Field<int>("HakenRyokin").ToString(); }

                if (chingin == DBNull.Value)
                {
                    TB_Chingin.Text = " ";
                }
                else { TB_Chingin.Text = dt.Rows[0].Field<int>("Chingin").ToString(); }

                if (tsukinteate == DBNull.Value)
                {
                    TB_TsukinTeate.Text = " ";
                }
                else { TB_TsukinTeate.Text = dt.Rows[0].Field<int>("TotalMoneyTrans").ToString(); }

                if (kyuyokojogaku == DBNull.Value)
                {
                    TB_KyuyoKojoGaku.Text = " ";
                }
                else { TB_KyuyoKojoGaku.Text = dt.Rows[0].Field<int>("KyuyoKojoGaku").ToString(); }

                if (worktime == DBNull.Value)
                {
                    TB_WorkTime.Text = " ";
                }
                else { TB_WorkTime.Text = dt.Rows[0].Field<int>("WorkTime").ToString(); }

                if (companyinsuredate == DBNull.Value || row["Shakaihoken"].ToString() == " ")
                {
                    DTP_CompanyInsureDate.Format = DateTimePickerFormat.Custom;
                    DTP_CompanyInsureDate.CustomFormat = " ";
                }
                else { DTP_CompanyInsureDate.Text = dt.Rows[0].Field<string>("Shakaihoken"); }

                if (koyohokendate == DBNull.Value || row["Kouyouhoken"].ToString() == " ")
                {
                    DTP_KoyoHokenDate.Format = DateTimePickerFormat.Custom;
                    DTP_KoyoHokenDate.CustomFormat = " ";
                }
                else { DTP_KoyoHokenDate.Text = dt.Rows[0].Field<string>("Kouyouhoken"); }

                if (dependentpeople == DBNull.Value)
                {
                    TB_DependentPeople.Text = " ";
                }
                else { TB_DependentPeople.Text = dt.Rows[0].Field<int>("DependentPeople").ToString(); }

                if (residentpeople == DBNull.Value)
                {
                    TB_ResidentPeople.Text = " ";
                }
                else { TB_ResidentPeople.Text = dt.Rows[0].Field<int>("ResidentPeople").ToString(); }

                if (healthinsurancepeople == DBNull.Value)
                {
                    TB_HealthInsurancePeople.Text = " ";
                }
                else { TB_HealthInsurancePeople.Text = dt.Rows[0].Field<int>("HealthInsurancePeople").ToString(); }

                if (zipcode == DBNull.Value)
                {
                    TB_ZipCode.Text = " ";
                }
                else { TB_ZipCode.Text = dt.Rows[0].Field<int>("ZipCode").ToString(); }

                if (zipcode1 == DBNull.Value)
                {
                    TB_ZipCode1.Text = " ";
                }
                else { TB_ZipCode1.Text = dt.Rows[0].Field<int>("ZipCode").ToString(); }
            }
        }

        // Cancel button click
        private void btCancel1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        ////////////////////////////////////////GET DATA TO ADD TO DATABASE/////////////////////////////////////

        public DTO_Edit updateData()
        {
            string _IDCode = TB_IDCode.Text;
            string _RomajiName = TB_RomajiName.Text;
            string _FuriganaName = TB_FuriganaName.Text;
            string _Address1 = TB_Address1.Text;
            string _Address3 = TB_Address3.Text;
            string _Address5 = TB_Address5.Text;
            string _CompanyCode = TB_CompanyCode.Text;
            string _CompanyName = TB_CompanyName.Text;
            string _BankName = TB_BankName.Text;
            string _BranchName = TB_BranchName.Text;
            string _AccountName = TB_AccountName.Text;
            string _BankCode = TB_BankCode.Text;
            string _BranchCode = TB_BranchCode.Text;
            // string _TeateGaku = TB_BranchCode.Text;
            string _AccountCode = TB_AccountCode.Text;
            string _AccountCode1 = TB_AccountCode1.Text;
            string _AccountCode2 = TB_AccountCode2.Text;
            string _AccountCode3 = TB_AccountCode3.Text;
            string _AccountCode4 = TB_AccountCode4.Text;
            string _AccountCode5 = TB_AccountCode5.Text;
            string _AccountCode6 = TB_AccountCode6.Text;
            string _AccountCode7 = TB_AccountCode7.Text;
            string _CompanyInsureDate = CheckDateTime(DTP_CompanyInsureDate);
            string _KoyoHokenDate = CheckDateTime(DTP_KoyoHokenDate);
            string _Birth = CheckDateTime(DTP_Birth); ;
            string _Reason = TB_Reason.Text;
            string _ChangeDate = CheckDateTime(DTP_ChangeDate);
            string _ChangeDateFrom = CheckDateTime(DTP_ChangeDateFrom);
            string _EmployTime1 = CheckDateTime(DTP_EmployTime1);
            string _EmployTime2 = CheckDateTime(DTP_EmployTime2);
            string _CardTimeOver = CheckDateTime(DTP_CardTimeOver);
            string _CardTimeStart = CheckDateTime(DTP_CardTimeStart);

            int _ZipCode;
            bool result_zipcode = int.TryParse(TB_ZipCode.Text, out _ZipCode);
            int _HakenRyokin;
            bool result_HakenRyokin = int.TryParse(TB_HakenRyokin.Text, out _HakenRyokin);
            int _Chingin;
            bool result_Chingin = int.TryParse(TB_Chingin.Text, out _Chingin);
            int _TsukinTeate;
            bool result_TsukinTeate = int.TryParse(TB_TsukinTeate.Text, out _TsukinTeate);
            int _KyuyoKojoGaku;
            bool result_KyuyoKojoGaku = int.TryParse(TB_KyuyoKojoGaku.Text, out _KyuyoKojoGaku);
            int _WorkTime;
            bool result_WorkTime = int.TryParse(TB_WorkTime.Text, out _WorkTime);
            int _DependentPeople;
            bool result_DependentPeople = int.TryParse(TB_DependentPeople.Text, out _DependentPeople);
            int _ResidentPeople;
            bool result_ResidentPeople = int.TryParse(TB_ResidentPeople.Text, out _ResidentPeople);
            int _HealthInsurancePeople;
            bool result_HealthInsurancePeople = int.TryParse(TB_HealthInsurancePeople.Text, out _HealthInsurancePeople);
            float _Genkaritsu;
            bool result_Genkaritsu = float.TryParse("", out _Genkaritsu);


            string _WorkType, _Tax, _ShiharaiType, _Sex, _TravelType, _CardType, _ClosingDate, _HakenRyokinType,
                 _ChinginType, _TeateType, _BankNameType, _BranchNameType, _Address2, _Address4;
            if (CB_Address2.SelectedIndex != -1) { _Address2 = CB_Address2.SelectedItem.ToString(); } else { _Address2 = string.Empty; }
            if (CB_Address4.SelectedIndex != -1) { _Address4 = CB_Address4.SelectedItem.ToString(); } else { _Address4 = string.Empty; }
            if (CB_WorkType.SelectedIndex != -1) { _WorkType = CB_WorkType.SelectedItem.ToString(); } else { _WorkType = string.Empty; }
            if (CB_Tax.SelectedIndex != -1) { _Tax = CB_Tax.SelectedItem.ToString(); } else { _Tax = string.Empty; }
            if (CB_ShiharaiType.SelectedIndex != -1) { _ShiharaiType = CB_ShiharaiType.SelectedItem.ToString(); } else { _ShiharaiType = string.Empty; }
            if (CB_Sex.SelectedIndex != -1) { _Sex = CB_Sex.SelectedItem.ToString(); } else { _Sex = string.Empty; }
            if (CB_TravelType.SelectedIndex != -1) { _TravelType = CB_TravelType.SelectedItem.ToString(); } else { _TravelType = string.Empty; }
            if (CB_CardType.SelectedIndex != -1) { _CardType = CB_CardType.SelectedItem.ToString(); } else { _CardType = string.Empty; }
            if (CB_ClosingDate.SelectedIndex != -1) { _ClosingDate = CB_ClosingDate.SelectedItem.ToString(); } else { _ClosingDate = string.Empty; }
            if (CB_HakenRyokinType.SelectedIndex != -1) { _HakenRyokinType = CB_HakenRyokinType.SelectedItem.ToString(); } else { _HakenRyokinType = string.Empty; }
            if (CB_ChinginType.SelectedIndex != -1) { _ChinginType = CB_ChinginType.SelectedItem.ToString(); } else { _ChinginType = string.Empty; }
            if (CB_TeateType.SelectedIndex != -1) { _TeateType = CB_TeateType.SelectedItem.ToString(); } else { _TeateType = string.Empty; }
            if (CB_BankNameType.SelectedIndex != -1) { _BankNameType = CB_BankNameType.SelectedItem.ToString(); } else { _BankNameType = string.Empty; }
            if (CB_BranchNameType.SelectedIndex != -1) { _BranchNameType = CB_BranchNameType.SelectedItem.ToString(); } else { _BranchNameType = string.Empty; }


            int _SeikinTeate = old_SeikinTeate, _GaikinTeate = old_GaikinTeate, _GijutsuTeate = old_GijutsuTeate,
                _ShikakuTeate = old_ShikakuTeate, _YakushokuTeate = old_YakushokuTeate, _EigyoTeate = old_EigyoTeate,
                _JutakuTeate = old_JutakuTeate, _BekkyoTeate = old_BekkyoTeate, _KazokuTeate = old_KazokuTeate;
            if (CB_TeateType.SelectedIndex > -1)
            {
                switch (this.CB_TeateType.SelectedItem.ToString())
                {
                    case "精勤手当":
                        bool result_SeikinTeate = int.TryParse(TB_TeateGaku.Text, out _SeikinTeate);
                        break;
                    case "外勤手当":
                        bool result_GaikinTeate = int.TryParse(TB_TeateGaku.Text, out _GaikinTeate);
                        break;
                    case "技術手当":
                        bool result_GijutsuTeate = int.TryParse(TB_TeateGaku.Text, out _GijutsuTeate);
                        break;
                    case "資格手当":
                        bool result_ShikakuTeate = int.TryParse(TB_TeateGaku.Text, out _ShikakuTeate);
                        break;
                    case "役職手当":
                        bool result_YakushokuTeate = int.TryParse(TB_TeateGaku.Text, out _YakushokuTeate);
                        break;
                    case "営業・職務手当":
                        bool result_EigyoTeate = int.TryParse(TB_TeateGaku.Text, out _EigyoTeate);
                        break;
                    case "家族手当":
                        bool result_KazokuTeate = int.TryParse(TB_TeateGaku.Text, out _KazokuTeate);
                        break;
                    case "住宅手当":
                        bool result_JutakuTeate = int.TryParse(TB_TeateGaku.Text, out _JutakuTeate);
                        break;
                    case "別居手当":
                        bool result_BekkyoTeate = int.TryParse(TB_TeateGaku.Text, out _BekkyoTeate);
                        break;
                }
            }
            DTO_Edit dto_edit = new DTO_Edit(_RomajiName, _IDCode, _FuriganaName, _CompanyName,
             _CompanyCode, _Sex, _ShiharaiType, _Tax, _Birth, _Reason,
             _ChangeDate, _ChangeDateFrom, _ZipCode, _Address1, _Address2,
             _Address3, _Address4, _Address5, _TravelType,
             _EmployTime1, _EmployTime2, _CardType, _CardTimeOver,
             _CardTimeStart, _WorkType, _ClosingDate, _HakenRyokin, _ChinginType,
             _HakenRyokinType, _Chingin, _TsukinTeate, _TeateType, _Genkaritsu,
             _KyuyoKojoGaku, _WorkTime, _BankName, _BankNameType,
             _BranchName, _BranchNameType, _AccountName, _BankCode,
             _BranchCode, _AccountCode, _AccountCode1, _AccountCode2, _AccountCode3,
             _AccountCode4, _AccountCode5, _AccountCode6, _AccountCode7, _CompanyInsureDate, _KoyoHokenDate,
             _DependentPeople, _ResidentPeople, _HealthInsurancePeople, _SeikinTeate, _GaikinTeate, _GijutsuTeate, _ShikakuTeate,
             _YakushokuTeate, _JutakuTeate, _BekkyoTeate, _KazokuTeate, _EigyoTeate);
            return dto_edit;
        }

        ///////////////////////////////////// CLICK SAVE BUTTON/////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////
        private void btSave_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("保存を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_edit.Update(updateData()))
                {
                    MessageBox.Show("保存しました。");
                    GUI_Main obj = (GUI_Main)Application.OpenForms["GUI_Main"];
                    obj.LoadGridView();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("保存は失敗しました。");
                }

            }
            else if (dialogResult == DialogResult.No)
            {
            }
        }

        public string CheckDateTime(DateTimePicker dtp)
        {
            string d;
            if (dtp.Text.Trim() == string.Empty)
            {
                d = " ";
            }
            else
            {
                d = bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtp);
            }
            return d;
        }

        /////////////////////////////////EVENT HIEN THI NGAY THANG NHU DUOCJ CHON TREN GUI///////////////////////////
        private void DTP_EmployTime1_ValueChanged(object sender, EventArgs e)
        {
            DTP_EmployTime1.Format = DateTimePickerFormat.Long;
        }

        private void DTP_Birth_ValueChanged(object sender, EventArgs e)
        {
            DTP_Birth.Format = DateTimePickerFormat.Long;
        }

        private void DTP_ChangeDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_ChangeDate.Format = DateTimePickerFormat.Long;
        }

        private void DTP_ChangeDateFrom_ValueChanged(object sender, EventArgs e)
        {
            DTP_ChangeDateFrom.Format = DateTimePickerFormat.Long;
        }

        private void DTP_EmployTime2_ValueChanged(object sender, EventArgs e)
        {
            DTP_EmployTime2.Format = DateTimePickerFormat.Long;
        }

        private void DTP_CardTimeStart_ValueChanged(object sender, EventArgs e)
        {
            DTP_CardTimeStart.Format = DateTimePickerFormat.Long;
        }

        private void DTP_CardTimeOver_ValueChanged(object sender, EventArgs e)
        {
            DTP_CardTimeOver.Format = DateTimePickerFormat.Long;
        }

        ///////////////////////EVENT TSUKINTEATE CLICK/////////////////////////////////////////
        private void TB_TsukinTeate_Click(object sender, EventArgs e)
        {
            GUI_Travel gui_travel = new GUI_Travel(name);
            gui_travel.Show();
        }

        /////////////////////////////////////EVENT ONLY INPUT NUMBER//////////////////////////////////////////
        private void TB_DependentPeople_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_ResidentPeople_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_HealthInsurancePeople_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_ZipCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_ZipCode1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_HakenRyokin_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_Chingin_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_TeateGaku_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_KyuyoKojoGaku_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_WorkTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }


        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x84:
                    base.WndProc(ref m);
                    if ((int)m.Result == 0x1)
                        m.Result = (IntPtr)0x2;
                    return;
            }

            base.WndProc(ref m);
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel3.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
        }
        private void panel4_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel4.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel5.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel6.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel7.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
        }

       
        private void TB_DependentPeople_Click(object sender, EventArgs e)
        {
            GUI_Dependent obj = new GUI_Dependent(name);
            obj.Show();
        }

        private void TB_AccountCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void TB_AccountCode7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }
       
        ////////////////////////////////EVENT XU LY SU KIEN CHON CB DE HIEN THI DUNG SO TIEN TEATEGAKU///////////////
        private void CB_TeateType_SelectedValueChanged(object sender, EventArgs e)
        {
            DataTable dt1 = bll_edit.EditForm(name);
            foreach (DataRow row in dt1.Rows)
            {
                if (CB_TeateType.SelectedIndex > -1)
                {
                    switch (this.CB_TeateType.SelectedItem.ToString())
                    {
                        case "精勤手当":
                            if (string.IsNullOrEmpty(row["SeikinTeate"].ToString()) || row["SeikinTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("SeikinTeate").ToString(); }
                            break;
                        case "外勤手当":
                            if (string.IsNullOrEmpty(row["GaikinTeate"].ToString()) || row["GaikinTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("GaikinTeate").ToString(); }
                            break;
                        case "技術手当":
                            if (string.IsNullOrEmpty(row["GijutsuTeate"].ToString()) || row["GijutsuTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("GijutsuTeate").ToString(); }

                            break;
                        case "資格手当":
                            if (string.IsNullOrEmpty(row["ShikakuTeate"].ToString()) || row["ShikakuTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("ShikakuTeate").ToString(); }
                            break;
                        case "役職手当":
                            if (string.IsNullOrEmpty(row["YakushokuTeate"].ToString()) || row["YakushokuTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("YakushokuTeate").ToString(); }

                            break;
                        case "営業・職務手当":
                            if (string.IsNullOrEmpty(row["EigyoTeate"].ToString()) || row["EigyoTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("EigyoTeate").ToString(); }
                            break;
                        case "家族手当":
                            if (string.IsNullOrEmpty(row["KazokuTeate"].ToString()) || row["KazokuTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("KazokuTeate").ToString(); }
                            break;
                        case "住宅手当":
                            if (string.IsNullOrEmpty(row["JutakuTeate"].ToString()) || row["JutakuTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("JutakuTeate").ToString(); }
                            break;
                        case "別居手当":
                            if (string.IsNullOrEmpty(row["BekkyoTeate"].ToString()) || row["BekkyoTeate"].ToString() == " ")
                            {
                                TB_TeateGaku.Text = " ";
                            }
                            else { TB_TeateGaku.Text = dt1.Rows[0].Field<int>("BekkyoTeate").ToString(); }
                            break;
                    }
                }
            }

        }

        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook = null;
        Excel.Worksheet xlWorkSheet = null;

        private void button1_Click(object sender, EventArgs e)
        {
            //DialogResult dialogResult = MessageBox.Show("印刷して保存します。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            //if (dialogResult == DialogResult.Yes)
            //{
            //    if (bll_edit.Update(updateData()))
            //    {
            //        MessageBox.Show("保存しました。");
            //        GUI_Main obj = (GUI_Main)Application.OpenForms["GUI_Main"];
            //        obj.LoadGridView();
            //        this.Close();
            //    }
            //    else
            //    {
            //        MessageBox.Show("保存は失敗しました。");
            //    }

            //}
            //else if (dialogResult == DialogResult.No)
            //{
            //}

            Print();
        }

        /////////////////////////////////HAM XU LY IN RA FILE VA EXXPORT RA EXCEL////////////////////////////////////
        private void Print()
        {
            DataTable dt = bll_edit.EditForm(name);
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(path + @"\File\template.xls",
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);

                xlWorkSheet.Cells[4, "BR"] = dt.Rows[0].Field<string>("Position");
                xlWorkSheet.Cells[7, "BR"] = dt.Rows[0].Field<string>("CreatePeople");
                xlWorkSheet.Cells[14, "E"] = dt.Rows[0].Field<string>("IDCode");
                xlWorkSheet.Cells[15, "U"] = dt.Rows[0].Field<string>("RomajiName");
                xlWorkSheet.Cells[14, "U"] = dt.Rows[0].Field<string>("FuriganaName");
                xlWorkSheet.Cells[15, "BD"] = dt.Rows[0].Field<string>("Sex");
                xlWorkSheet.Cells[15, "BJ"] = dt.Rows[0].Field<string>("ShiharaiType");
                xlWorkSheet.Cells[15, "BP"] = dt.Rows[0].Field<string>("ClosingDate");
                if (dt.Rows[0].Field<string>("Birth") != " ")
                {
                    xlWorkSheet.Cells[15, "BV"] = dt.Rows[0].Field<string>("Birth");
                }
                xlWorkSheet.Cells[17, "U"] = dt.Rows[0].Field<string>("CompanyName");
                xlWorkSheet.Cells[17, "BJ"] = TB_Reason.Text;

                string temp_ChangeDate = CheckDateTime(DTP_ChangeDate);
                if (temp_ChangeDate != " ")
                {
                    string[] changedate = temp_ChangeDate.Split('/');
                    xlWorkSheet.Cells[21, "W"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[21, "AF"] = changedate[1];
                    xlWorkSheet.Cells[21, "AP"] = changedate[2];
                }
                string temp_ChangeDateFrom = CheckDateTime(DTP_ChangeDateFrom);
                if (temp_ChangeDateFrom != " ")
                {
                    string[] changedateFrom = temp_ChangeDateFrom.Split('/');
                    xlWorkSheet.Cells[24, "W"] = (Convert.ToInt32(changedateFrom[0]) - 1988).ToString();
                    xlWorkSheet.Cells[24, "AF"] = changedateFrom[1];
                    xlWorkSheet.Cells[24, "AP"] = CB_TravelType.SelectedItem.ToString();
                }
                /////////////In cả hai bảng những chỗ thay đổi///////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////////////////

                if (TB_RomajiName.Text != dt.Rows[0].Field<string>("RomajiName"))
                {
                    xlWorkSheet.Cells[30, "P"] = dt.Rows[0].Field<string>("RomajiName");
                    xlWorkSheet.Cells[30, "BB"] = TB_RomajiName.Text;
                    xlWorkSheet.Cells[30, "E"] = "☑氏　名";
                    xlWorkSheet.Cells[30, "E"].Font.Bold = true;
                }
                if (dt.Rows[0].Field<string>("CardType") != CB_TravelType.SelectedItem.ToString())
                {
                    xlWorkSheet.Cells[45, "E"] = "☑在留資格";
                    xlWorkSheet.Cells[45, "E"].Font.Bold = true;
                    //Ben trai
                    switch (dt.Rows[0].Field<string>("CardType"))
                    {
                        case "定住者":
                            xlWorkSheet.Cells[45, "P"] = "☑定・永・特永・日配・永配・その他（";
                            break;
                        case "永住者":
                            xlWorkSheet.Cells[45, "P"] = "定・☑永・特永・日配・永配・その他（";
                            break;
                        case "特別永住":
                            xlWorkSheet.Cells[45, "P"] = "定・☑永・☑特永・日配・永配・その他（";
                            break;
                        case "日本人配":
                            xlWorkSheet.Cells[45, "P"] = "定・☑永・特永・☑日配・永配・その他（";
                            break;
                        case "永住配":
                            xlWorkSheet.Cells[45, "P"] = "定・☑永・特永・日配・☑永配・その他（";
                            break;
                        default:
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・日配・☑永配・☑その他（";
                            xlWorkSheet.Cells[45, "AQ"] = dt.Rows[0].Field<string>("CardType");
                            break;
                    }
                    //Ben phai
                    switch (CB_TravelType.SelectedItem.ToString())
                    {
                        case "定住者":
                            xlWorkSheet.Cells[45, "BB"] = "☑定・永・特永・日配・永配・その他（";
                            break;
                        case "永住者":
                            xlWorkSheet.Cells[45, "BB"] = "定・☑永・特永・日配・永配・その他（";
                            break;
                        case "特別永住":
                            xlWorkSheet.Cells[45, "BB"] = "定・☑永・☑特永・日配・永配・その他（";
                            break;
                        case "日本人配":
                            xlWorkSheet.Cells[45, "BB"] = "定・☑永・特永・☑日配・永配・その他（";
                            break;
                        case "永住配":
                            xlWorkSheet.Cells[45, "BB"] = "定・☑永・特永・日配・☑永配・その他（";
                            break;
                        default:
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・日配・☑永配・☑その他（";
                            xlWorkSheet.Cells[45, "CD"] = CB_TravelType.SelectedItem.ToString();
                            break;
                    }

                }
                if (dt.Rows[0].Field<string>("ClosingDate") != CB_ClosingDate.SelectedItem.ToString())
                {
                    xlWorkSheet.Cells[54, "E"] = "☑締日";
                    xlWorkSheet.Cells[54, "E"].Font.Bold = true;
                    //Ben trai
                    switch (dt.Rows[0].Field<string>("ClosingDate"))
                    {
                        case "15":
                            xlWorkSheet.Cells[54, "P"] = "☑１５　　・　　２０　　・　　２５　　・　　末";
                            break;
                        case "20":
                            xlWorkSheet.Cells[54, "P"] = "１５　　・　　☑２０　　・　　２５　　・　　末";
                            break;
                        case "25":
                            xlWorkSheet.Cells[54, "P"] = "１５　　・　　２０　　・　　☑２５　　・　　末";
                            break;
                        case "末":
                            xlWorkSheet.Cells[54, "P"] = "１５　　・　　２０　　・　　２５　　・　　☑末";
                            break;
                    }
                    // Ben phai
                    switch (CB_ClosingDate.SelectedItem.ToString())
                    {
                        case "15":
                            xlWorkSheet.Cells[54, "BB"] = "☑１５　　・　　２０　　・　　２５　　・　　末";
                            break;
                        case "20":
                            xlWorkSheet.Cells[54, "BB"] = "１５　　・　　☑２０　　・　　２５　　・　　末";
                            break;
                        case "25":
                            xlWorkSheet.Cells[54, "BB"] = "１５　　・　　２０　　・　　☑２５　　・　　末";
                            break;
                        case "末":
                            xlWorkSheet.Cells[54, "BB"] = "１５　　・　　２０　　・　　２５　　・　　☑末";
                            break;
                    }
                }

                //if (dt.Rows[0].Field<string>("WorkType") != CB_WorkType.SelectedItem.ToString())
                //{
                //    xlWorkSheet.Cells[54, "E"] = "☑就労形態";
                //    xlWorkSheet.Cells[54, "E"].Font.Bold = true;
                //    //Ben trai
                //    switch (dt.Rows[0].Field<string>("WorkType"))
                //    {
                //        case "請負":
                //            xlWorkSheet.Cells[57, "P"] = "派遣　　・　　☑請負";
                //            break;
                //        case "派遣":
                //            xlWorkSheet.Cells[57, "P"] = "☑派遣　　・　　請負";
                //            break;
                //    }
                //    // Ben phai
                //    switch (CB_WorkType.SelectedItem.ToString())
                //    {
                //        case "請負":
                //            xlWorkSheet.Cells[57, "BB"] = "派遣　　・　　☑請負";
                //            break;
                //        case "派遣":
                //            xlWorkSheet.Cells[57, "BB"] = "☑派遣　　・　　請負";
                //            break;
                //    }
                //}
                //string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                //if (zipcode != TB_ZipCode.Text || TB_Address1.Text != dt.Rows[0].Field<string>("Address1") ||
                //    CB_Address2.SelectedItem.ToString() != dt.Rows[0].Field<string>("Address2") || TB_Address3.Text != dt.Rows[0].Field<string>("Address3")
                //    || CB_Address4.SelectedItem.ToString() != dt.Rows[0].Field<string>("Address4") || TB_Address5.Text != dt.Rows[0].Field<string>("Address5"))
                //{
                //    //Bảng bên phải
                //    xlWorkSheet.Cells[33, "E"] = "☑住民票住所";
                //    xlWorkSheet.Cells[33, "E"].Font.Bold = true;
                //    xlWorkSheet.Cells[33, "BE"] = TB_ZipCode.Text;
                //    // xlWorkSheet.Cells[36, "BE"] = TB_ZipCode.Text;
                //    string address2 = TB_Address1.Text + CB_Address2.SelectedItem.ToString() +
                //    TB_Address3.Text + CB_Address4.SelectedItem.ToString() + TB_Address5.Text;
                //    xlWorkSheet.Cells[33, "BK"] = address2;
                //    //// Bang ben trái
                //    xlWorkSheet.Cells[33, "S"] = zipcode;
                //    xlWorkSheet.Cells[36, "S"] = zipcode;
                //}

                //if (dt.Rows[0].Field<string>("TravelType") != CB_TravelType.SelectedItem.ToString())
                //{   ///Bảng bên phải
                //    xlWorkSheet.Cells[39, "E"] = "☑通勤/入寮";
                //    xlWorkSheet.Cells[39, "E"].Font.Bold = true;
                //    if (CB_TravelType.SelectedItem.ToString() == "入寮")
                //    {
                //        xlWorkSheet.Cells[39, "BK"] = "☑";
                //    }
                //    else
                //    {
                //        xlWorkSheet.Cells[39, "BB"] = "☑";
                //    }
                //    // Bảng bên trái
                //    if (dt.Rows[0].Field<string>("TravelType") == "入寮")
                //    {
                //        xlWorkSheet.Cells[39, "Y"] = "☑";
                //    }
                //    else
                //    {
                //        xlWorkSheet.Cells[39, "P"] = "☑";
                //    }

                //}

                //if (dt.Rows[0].Field<string>("EmployTime1") != CheckDateTime(DTP_EmployTime1) || dt.Rows[0].Field<string>("EmployTime2") != CheckDateTime(DTP_EmployTime2))
                //{   //Bảng bên phải
                //    string temp_time1 = CheckDateTime(DTP_EmployTime1);
                //    string[] Time1_temps = temp_time1.Split('/');
                //    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                //    string temp_time2 = CheckDateTime(DTP_EmployTime2);
                //    string[] Time2_temps = temp_time2.Split('/');
                //    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                //    xlWorkSheet.Cells[42, "BB"] = a + b;
                //    xlWorkSheet.Cells[42, "E"] = "☑雇用期間";
                //    xlWorkSheet.Cells[42, "E"].Font.Bold = true;

                //    //Bảng bên trái
                //    if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
                //    {
                //        string temp_time11 = dt.Rows[0].Field<string>("EmployTime1");
                //        string[] Time11_temps = temp_time11.Split('/');
                //        string a1 = "平成 " + (Convert.ToInt32(Time11_temps[0]) - 1988).ToString() + " 年 " + Time11_temps[1] + " 月 " + Time11_temps[2] + " 日から ";

                //        string temp_time21 = dt.Rows[0].Field<string>("EmployTime2");
                //        string[] Time21_temps = temp_time21.Split('/');
                //        string b1 = "平成 " + (Convert.ToInt32(Time21_temps[0]) - 1988).ToString() + " 年 " + Time21_temps[1] + " 月 " + Time21_temps[2] + " 日まで";

                //        xlWorkSheet.Cells[42, "P"] = a1 + b1;
                //    }
                //}

                //if (dt.Rows[0].Field<string>("CardTime") != CheckDateTime(DTP_CardTimeStart) || dt.Rows[0].Field<string>("CardTimeOut") != CheckDateTime(DTP_CardTimeOver))
                //{   // Bên phải
                //    string temp_time1 = CheckDateTime(DTP_CardTimeStart);
                //    string[] Time1_temps = temp_time1.Split('/');
                //    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                //    string temp_time2 = CheckDateTime(DTP_CardTimeOver);
                //    string[] Time2_temps = temp_time2.Split('/');
                //    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                //    xlWorkSheet.Cells[48, "BB"] = a + b;
                //    xlWorkSheet.Cells[48, "E"] = "☑在留期間";
                //    xlWorkSheet.Cells[48, "E"].Font.Bold = true;

                //    //Bên trái
                //    if (dt.Rows[0].Field<string>("Nationality") != string.Empty)
                //    {
                //        string temp_time11 = dt.Rows[0].Field<string>("CardTime");
                //        string[] Time11_temps = temp_time11.Split('/');
                //        string a1 = "平成 " + (Convert.ToInt32(Time11_temps[0]) - 1988).ToString() + " 年 " + Time11_temps[1] + " 月 " + Time11_temps[2] + " 日から ";

                //        string temp_time21 = dt.Rows[0].Field<string>("CardTimeOut");
                //        string[] Time21_temps = temp_time21.Split('/');
                //        string b1 = "平成 " + (Convert.ToInt32(Time21_temps[0]) - 1988).ToString() + " 年 " + Time21_temps[1] + " 月 " + Time21_temps[2] + " 日まで";

                //        xlWorkSheet.Cells[48, "P"] = a1 + b1;
                //    }
                //}

                //if (TB_CompanyName.Text != dt.Rows[0].Field<string>("CompanyName"))
                //{   //Bên phải
                //    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                //    xlWorkSheet.Cells[51, "E"] = "☑企業名";
                //    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                //    xlWorkSheet.Cells[51, "BB"] = TB_CompanyName.Text;
                //    //Bên trái
                //    xlWorkSheet.Cells[51, "P"] = dt.Rows[0].Field<string>("CompanyName");
                //}

                //if (CB_ShiharaiType.SelectedItem.ToString() != dt.Rows[0].Field<string>("ShiharaiType"))
                //{
                //    xlWorkSheet.Cells[60, "BF"] = CB_ShiharaiType.SelectedItem.ToString();
                //    xlWorkSheet.Cells[60, "E"] = "☑賃金支払形態";
                //    xlWorkSheet.Cells[60, "E"].Font.Bold = true;
                //    //Bên trái
                //    xlWorkSheet.Cells[60, "T"] = dt.Rows[0].Field<string>("ShiharaiType");
                //}

                //if (CB_Tax.SelectedItem.ToString() != dt.Rows[0].Field<string>("Tax"))
                //{   //Bên phải
                //    xlWorkSheet.Cells[60, "CA"] = CB_Tax.SelectedItem.ToString();
                //    xlWorkSheet.Cells[62, "E"] = "☑税適";
                //    xlWorkSheet.Cells[62, "E"].Font.Bold = true;
                //    // Bên trái
                //    xlWorkSheet.Cells[60, "AM"] = dt.Rows[0].Field<string>("Tax");
                //}

                //if (TB_HakenRyokin.Text != dt.Rows[0].Field<int?>("HakenRyokin").ToString() || CB_HakenRyokinType.SelectedItem.ToString() != dt.Rows[0].Field<string>("HakenRyokinType").ToString())
                //{   //Bên phải
                //    xlWorkSheet.Cells[67, "BF"] = TB_HakenRyokin.Text;
                //    xlWorkSheet.Cells[67, "BT"] = CB_HakenRyokinType.SelectedItem.ToString();
                //    xlWorkSheet.Cells[67, "E"] = "☑派遣・請金";
                //    xlWorkSheet.Cells[67, "E"].Font.Bold = true;
                //    //Bên trái
                //    xlWorkSheet.Cells[67, "T"] = dt.Rows[0].Field<int?>("HakenRyokin");
                //    xlWorkSheet.Cells[67, "AX"] = dt.Rows[0].Field<string>("HakenRyokinType");
                //}
                //if (TB_KyuyoKojoGaku.Text != dt.Rows[0].Field<int?>("DormitoryFee").ToString())
                //{
                //    xlWorkSheet.Cells[85, "AJ"] = dt.Rows[0].Field<int?>("DormitoryFee");
                //    xlWorkSheet.Cells[85, "R"] = "寮費";

                //    xlWorkSheet.Cells[85, "BO"] = TB_KyuyoKojoGaku.Text;
                //    xlWorkSheet.Cells[85, "BD"] = "寮費";
                //    xlWorkSheet.Cells[85, "E"] = "☑給与控除額";
                //    xlWorkSheet.Cells[85, "E"].Font.Bold = true;

                //}

                //if (TB_Chingin.Text != dt.Rows[0].Field<int?>("Chingin").ToString() || CB_ChinginType.SelectedItem.ToString() != dt.Rows[0].Field<string>("ChinginType").ToString())
                //{   //Bên phải
                //    xlWorkSheet.Cells[70, "BF"] = TB_Chingin.Text;
                //    xlWorkSheet.Cells[70, "BT"] = dt.Rows[0].Field<string>("ChinginType");
                //    xlWorkSheet.Cells[70, "E"] = "☑賃金";
                //    xlWorkSheet.Cells[70, "E"].Font.Bold = true;
                //    //Bên trái
                //    xlWorkSheet.Cells[70, "T"] = dt.Rows[0].Field<int?>("Chingin");
                //    xlWorkSheet.Cells[70, "AX"] = dt.Rows[0].Field<string>("ChinginType");
                //}

                //if (old_tsukinTeate != TB_TsukinTeate.Text)
                //{   // Bên phải
                //    xlWorkSheet.Cells[76, "BB"] = TB_TsukinTeate.Text;
                //    xlWorkSheet.Cells[76, "E"] = "☑通勤手当";
                //    xlWorkSheet.Cells[76, "E"].Font.Bold = true;
                //    // bên trái
                //    xlWorkSheet.Cells[76, "P"] = old_tsukinTeate;
                //}

                //if (TB_BankCode.Text != dt.Rows[0].Field<string>("BankCode") || TB_BranchCode.Text != dt.Rows[0].Field<string>("BranchCode")
                //    || TB_BankName.Text != dt.Rows[0].Field<string>("BankName") || CB_BankNameType.SelectedItem.ToString() != dt.Rows[0].Field<string>("BankNameType") ||
                //    TB_BranchName.Text != dt.Rows[0].Field<string>("BranchName") || CB_BranchNameType.SelectedItem.ToString() != dt.Rows[0].Field<string>("BranchNameType"))
                //{   // Bên phải
                //    xlWorkSheet.Cells[91, "E"] = "☑銀行/支店名";
                //    xlWorkSheet.Cells[91, "E"].Font.Bold = true;
                //    xlWorkSheet.Cells[91, "BH"] = TB_BankCode.Text;
                //    xlWorkSheet.Cells[91, "CC"] = TB_BranchCode.Text;
                //    xlWorkSheet.Cells[92, "BB"] = TB_BankName.Text;
                //    xlWorkSheet.Cells[92, "BT"] = CB_BankNameType.SelectedItem.ToString();
                //    xlWorkSheet.Cells[92, "BW"] = TB_BranchName.Text;
                //    xlWorkSheet.Cells[92, "CO"] = CB_BranchNameType.SelectedItem.ToString();
                //    // Bên trái(cũ)
                //    xlWorkSheet.Cells[91, "U"] = dt.Rows[0].Field<string>("BankCode");
                //    xlWorkSheet.Cells[91, "AN"] = dt.Rows[0].Field<string>("BranchCode");
                //    xlWorkSheet.Cells[92, "P"] = dt.Rows[0].Field<string>("BankName");
                //    xlWorkSheet.Cells[92, "AF"] = dt.Rows[0].Field<string>("BankNameType");
                //    xlWorkSheet.Cells[92, "AI"] = dt.Rows[0].Field<string>("BranchName");
                //    xlWorkSheet.Cells[92, "AY"] = dt.Rows[0].Field<string>("BranchNameType");

                //}
                //if (TB_AccountName.Text != dt.Rows[0].Field<string>("AccountName"))
                //{
                //    xlWorkSheet.Cells[95, "E"] = "☑口座名義（カナ";
                //    xlWorkSheet.Cells[95, "E"].Font.Bold = true;
                //    //Ben phai
                //    xlWorkSheet.Cells[95, "BB"] = TB_AccountName.Text;
                //    //Ben trai
                //    xlWorkSheet.Cells[95, "P"] = dt.Rows[0].Field<string>("AccountName");
                //}
                //if (TB_AccountCode.Text != dt.Rows[0].Field<string>("AccountCode1") || TB_AccountCode1.Text != dt.Rows[0].Field<string>("AccountCode2") ||
                //TB_AccountCode2.Text != dt.Rows[0].Field<string>("AccountCode3") || TB_AccountCode3.Text != dt.Rows[0].Field<string>("AccountCode4") ||
                //TB_AccountCode4.Text != dt.Rows[0].Field<string>("AccountCode5") || TB_AccountCode5.Text != dt.Rows[0].Field<string>("AccountCode6") ||
                //TB_AccountCode6.Text != dt.Rows[0].Field<string>("AccountCode7") || TB_AccountCode7.Text != dt.Rows[0].Field<string>("AccountCode8"))
                //{
                //    xlWorkSheet.Cells[98, "E"] = "☑口座番号";
                //    xlWorkSheet.Cells[98, "E"].Font.Bold = true;
                //    //Bên phải
                //    xlWorkSheet.Cells[98, "BF"] = TB_AccountCode.Text + TB_AccountCode1.Text + TB_AccountCode2.Text + TB_AccountCode3.Text
                //    + TB_AccountCode4.Text + TB_AccountCode5.Text + TB_AccountCode6.Text + TB_AccountCode7.Text;
                //    //Bên trái
                //    xlWorkSheet.Cells[98, "T"] = dt.Rows[0].Field<string>("AccountCode1") + dt.Rows[0].Field<string>("AccountCode2") +
                //   dt.Rows[0].Field<string>("AccountCode3") + dt.Rows[0].Field<string>("AccountCode4") + dt.Rows[0].Field<string>("AccountCode5")
                //   + dt.Rows[0].Field<string>("AccountCode6") + dt.Rows[0].Field<string>("AccountCode7") + dt.Rows[0].Field<string>("AccountCode8");
                //}

                //if (CheckDateTime(DTP_KoyoHokenDate) != " " && CheckDateTime(DTP_KoyoHokenDate) != dt.Rows[0].Field<string>("Kouyouhoken"))
                //{   //Bên phải
                //    xlWorkSheet.Cells[101, "E"] = "☑雇用保険";
                //    xlWorkSheet.Cells[101, "E"].Font.Bold = true;
                //    string temp = CheckDateTime(DTP_KoyoHokenDate);
                //    string[] temps = temp.Split('/');
                //    xlWorkSheet.Cells[101, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                //    xlWorkSheet.Cells[101, "CJ"] = temps[1];
                //    xlWorkSheet.Cells[101, "CN"] = temps[2];
                //    //Bên trái
                //    if (dt.Rows[0].Field<string>("Kouyouhoken") != " ")
                //    {
                //        string temp1 = dt.Rows[0].Field<string>("Kouyouhoken");

                //        string[] temps1 = temp1.Split('/');
                //        xlWorkSheet.Cells[101, "AJ"] = (Convert.ToInt32(temps1[0]) - 1988).ToString();
                //        xlWorkSheet.Cells[101, "AO"] = temps1[1];
                //        xlWorkSheet.Cells[101, "AT"] = temps1[2];
                //    }
                //}
                //if (CheckDateTime(DTP_CompanyInsureDate) != " " && CheckDateTime(DTP_KoyoHokenDate) != dt.Rows[0].Field<string>("Shakaihoken"))
                //{   //Bên phải
                //    xlWorkSheet.Cells[104, "E"] = "☑社会保険";
                //    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                //    string temp = CheckDateTime(DTP_CompanyInsureDate);
                //    string[] temps = temp.Split('/');
                //    xlWorkSheet.Cells[104, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                //    xlWorkSheet.Cells[104, "CJ"] = temps[1];
                //    xlWorkSheet.Cells[104, "CN"] = temps[2];
                //    //Bên trái
                //    if (dt.Rows[0].Field<string>("Shakaihoken") != " ")
                //    {
                //        string temp1 = dt.Rows[0].Field<string>("Shakaihoken");

                //        string[] temps1 = temp1.Split('/');
                //        xlWorkSheet.Cells[104, "AJ"] = (Convert.ToInt32(temps1[0]) - 1988).ToString();
                //        xlWorkSheet.Cells[104, "AO"] = temps1[1];
                //        xlWorkSheet.Cells[104, "AT"] = temps1[2];
                //    }
                //}

                //if (TB_DependentPeople.Text != old_DependentPeople ||
                //    TB_ResidentPeople.Text != old_ResidentPeople ||
                //    TB_HealthInsurancePeople.Text != old_HealthInsurancePeople)
                //{   //Bên phải
                //    xlWorkSheet.Cells[104, "E"] = "☑扶 養 人 数";
                //    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                //    xlWorkSheet.Cells[107, "BK"] = TB_DependentPeople.Text;
                //    xlWorkSheet.Cells[107, "BY"] = TB_ResidentPeople.Text;
                //    xlWorkSheet.Cells[107, "CM"] = TB_HealthInsurancePeople.Text;
                //    // Bên trái
                //    xlWorkSheet.Cells[107, "X"] = old_DependentPeople;
                //    xlWorkSheet.Cells[107, "AK"] = old_ResidentPeople;
                //    xlWorkSheet.Cells[107, "AW"] = old_HealthInsurancePeople;
                //}
                /////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////In ca 4 trang truoc nua///////////////////////



                //cho nay de xu ly may in default
                //var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
                //int printerIndex = 0;
                //foreach (String s in printers)
                //{
                //    if (s.Equals("白黒　SHARP MX-2650FN SPDL2-c"))
                //    {
                //        break;
                //    }
                //    printerIndex++;
                //}

                //// Print out 1 copy to the default printer:
                //xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //                     printers[printerIndex], Type.Missing, Type.Missing, Type.Missing);

                //// Cleanup:
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                //Marshal.FinalReleaseComObject(xlWorkSheet);

                //xlWorkBook.Close(false, Type.Missing, Type.Missing);
                //Marshal.FinalReleaseComObject(xlWorkBook);

                //xlApp.Quit();
                //Marshal.FinalReleaseComObject(xlApp);
                //MessageBox.Show("印刷準備完了");

                ////////// show promt to save file
                System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
                saveDlg.InitialDirectory = @"C:\";
                saveDlg.Filter = "Excel files (*.xls)|*.xls";
                saveDlg.FilterIndex = 0;
                saveDlg.RestoreDirectory = true;
                saveDlg.Title = "Export Excel File To";
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
                    this.Close();
                }
                else
                {
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                }
            }
            catch (Exception eSavePrint)
            {
                // Cleanup Memory
                xlWorkBook.Close(0);
                xlApp.Quit();
                MessageBox.Show(eSavePrint.Message, "エラー！印刷できません！");
            }
        }

        private void DTP_KoyoHokenDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_KoyoHokenDate.Format = DateTimePickerFormat.Long;
        }

        private void DTP_CompanyInsureDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_CompanyInsureDate.Format = DateTimePickerFormat.Long;
        }

        private void TB_ZipCode_TextChanged(object sender, EventArgs e)
        {
             bll_handleFunc.AutoShowAddress(TB_ZipCode, TB_Address1, CB_Address2, TB_Address3, CB_Address4, TB_Address5);
        }











    }
}
