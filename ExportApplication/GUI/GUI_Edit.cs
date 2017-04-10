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
using System.Threading;

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
            label12.Font = new Font(label2.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
            label12.ForeColor = System.Drawing.Color.Red;
            label10.Font = new Font(label2.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
            label10.ForeColor = System.Drawing.Color.Red;
            label11.Font = new Font(label2.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
            label11.ForeColor = System.Drawing.Color.Red;
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

        DataTable history = new DataTable();

        private void GUI_Edit_Load(object sender, EventArgs e)
        {
            DataTable dt = bll_edit.EditForm(name);
            history = dt.Clone();
            old_tsukinTeate = dt.Rows[0].Field<int?>("TotalMoneyTrans").ToString();
            old_DependentPeople = dt.Rows[0].Field<int?>("DependentPeople").ToString();
            old_ResidentPeople = dt.Rows[0].Field<int?>("ResidentPeople").ToString();
            old_HealthInsurancePeople = dt.Rows[0].Field<int?>("HealthInsurancePeople").ToString();
            foreach (DataRow row in dt.Rows)
            {
                history.ImportRow(dt.Rows[0]);

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
                object KyuyoKojoGaku = row["DormitoryFee"];
                if (KyuyoKojoGaku == DBNull.Value)
                {
                    TB_KyuyoKojoGaku.Text = " ";
                }
                else { TB_KyuyoKojoGaku.Text = dt.Rows[0].Field<int>("DormitoryFee").ToString(); }
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
            DialogResult dialogResult = MessageBox.Show("EXCELファイルへ出力して保存します。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_edit.Update(updateData()))
                {
                    Export();
                    
                }
                else
                {
                    MessageBox.Show("保存は失敗しました。");
                }

            }
            else if (dialogResult == DialogResult.No)
            {
            }
            //foreach (DataRow row in history.Rows)
            //{ MessageBox.Show(history.Rows[0].Field<string>("FuriganaName")); }
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
            DialogResult dialogResult = MessageBox.Show("印刷して保存します。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_edit.Update(updateData()))
                {
                    Print();
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

        /////////////////////////////////HAM XU LY IN RA FILE VA EXXPORT RA EXCEL////////////////////////////////////
        private void Print()
        {
            Thread t = new Thread(new ThreadStart(SplashScreen));
            t.Start();
            Thread.Sleep(5000);

            BLL_Print bll_print = new BLL_Print();
            BLL_HandleFunc bll_handle = new BLL_HandleFunc();
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

                string temp_ChangeDate = dt.Rows[0].Field<string>("ChangeDate");
                if (temp_ChangeDate != " ")
                {
                    string[] changedate = temp_ChangeDate.Split('/');
                    xlWorkSheet.Cells[21, "W"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[21, "AF"] = changedate[1];
                    xlWorkSheet.Cells[21, "AP"] = changedate[2];

                    xlWorkSheet.Cells[64, "BF"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[64, "BL"] = changedate[1];
                    xlWorkSheet.Cells[64, "BQ"] = changedate[2];
                }
                string temp_ChangeDateFrom = dt.Rows[0].Field<string>("ChangeDateFrom");
                if (temp_ChangeDateFrom != " ")
                {
                    string[] changedateFrom = temp_ChangeDateFrom.Split('/');
                    xlWorkSheet.Cells[24, "W"] = (Convert.ToInt32(changedateFrom[0]) - 1988).ToString();
                    xlWorkSheet.Cells[24, "AF"] = changedateFrom[1];
                    xlWorkSheet.Cells[24, "AP"] = dt.Rows[0].Field<string>("ClosingDate");

                }
                /////////////In cả hai bảng những chỗ thay đổi///////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////////////////

                foreach (DataRow row in dt.Rows)
                {
                    if (CB_TeateType.SelectedIndex > -1 || old_SeikinTeate != dt.Rows[0].Field<int>("SeikinTeate") ||
                        old_GaikinTeate != dt.Rows[0].Field<int>("GaikinTeate") || old_YakushokuTeate != dt.Rows[0].Field<int>("GijutsuTeate") ||
                         old_ShikakuTeate != dt.Rows[0].Field<int>("ShikakuTeate") || old_YakushokuTeate != dt.Rows[0].Field<int>("YakushokuTeate") ||
                        old_EigyoTeate != dt.Rows[0].Field<int>("EigyoTeate") || old_KazokuTeate != dt.Rows[0].Field<int>("KazokuTeate") ||
                        old_JutakuTeate != dt.Rows[0].Field<int>("JutakuTeate") || old_BekkyoTeate != dt.Rows[0].Field<int>("BekkyoTeate"))
                    {
                        switch (this.CB_TeateType.SelectedItem.ToString())
                        {
                            case "精勤手当":
                                // Ben trai
                                xlWorkSheet.Cells[79, "R"] = "精勤手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_SeikinTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;

                                // Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "精勤手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("SeikinTeate").ToString();
                                break;
                            case "外勤手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "外勤手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_GaikinTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                // Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "外勤手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("GaikinTeate").ToString();
                                break;
                            case "技術手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "技術手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_GijutsuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                // ben phai
                                xlWorkSheet.Cells[79, "BD"] = "技術手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("GijutsuTeate").ToString();
                                break;
                            case "資格手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "資格手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_ShikakuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "資格手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("ShikakuTeate").ToString();
                                break;
                            case "役職手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "役職手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_YakushokuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "役職手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("YakushokuTeate").ToString();
                                break;
                            case "営業・職務手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "営業・職務手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_EigyoTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "営業・職務手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("EigyoTeate").ToString();
                                break;
                            case "家族手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "家族手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_KazokuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "家族手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("KazokuTeate").ToString();
                                break;
                            case "住宅手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "住宅手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_JutakuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "住宅手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("JutakuTeate").ToString();
                                break;
                            case "別居手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "別居手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_BekkyoTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "別居手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("BekkyoTeate").ToString();
                                break;
                        }
                    }
                }
                if (history.Rows[0].Field<string>("RomajiName") != dt.Rows[0].Field<string>("RomajiName"))
                {
                    xlWorkSheet.Cells[30, "P"] = history.Rows[0].Field<string>("RomajiName");
                    xlWorkSheet.Cells[30, "BB"] = dt.Rows[0].Field<string>("RomajiName");
                    xlWorkSheet.Cells[30, "E"] = "☑氏　名";
                    xlWorkSheet.Cells[30, "E"].Font.Bold = true;
                }
                if (dt.Rows[0].Field<string>("CardType") != history.Rows[0].Field<string>("CardType"))
                {
                    xlWorkSheet.Cells[45, "E"] = "☑在留資格";
                    xlWorkSheet.Cells[45, "E"].Font.Bold = true;
                    //Ben trai
                    switch (history.Rows[0].Field<string>("CardType"))
                    {
                        case "定住者":
                            xlWorkSheet.Cells[45, "P"] = "☑定・永・特永・日配・永配・その他（";
                            break;
                        case "永住者":
                            xlWorkSheet.Cells[45, "P"] = "定・☑永・特永・日配・永配・その他（";
                            break;
                        case "特別永住":
                            xlWorkSheet.Cells[45, "P"] = "定・永・☑特永・日配・永配・その他（";
                            break;
                        case "日本人配":
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・☑日配・永配・その他（";
                            break;
                        case "永住配":
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・日配・☑永配・その他（";
                            break;
                        default:
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・日配・永配・☑その他（";
                            xlWorkSheet.Cells[45, "AQ"] = dt.Rows[0].Field<string>("CardType");
                            break;
                    }
                    //Ben phai
                    switch (dt.Rows[0].Field<string>("CardType"))
                    {
                        case "定住者":
                            xlWorkSheet.Cells[45, "BB"] = "☑定・永・特永・日配・永配・その他（";
                            break;
                        case "永住者":
                            xlWorkSheet.Cells[45, "BB"] = "定・☑永・特永・日配・永配・その他（";
                            break;
                        case "特別永住":
                            xlWorkSheet.Cells[45, "BB"] = "定・永・☑特永・日配・永配・その他（";
                            break;
                        case "日本人配":
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・☑日配・永配・その他（";
                            break;
                        case "永住配":
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・日配・☑永配・その他（";
                            break;
                        default:
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・日配・永配・☑その他（";
                            xlWorkSheet.Cells[45, "CD"] = CB_CardType.SelectedItem.ToString();
                            break;
                    }

                }
                if (dt.Rows[0].Field<string>("ClosingDate") != history.Rows[0].Field<string>("ClosingDate"))
                {
                    xlWorkSheet.Cells[54, "E"] = "☑締日";
                    xlWorkSheet.Cells[54, "E"].Font.Bold = true;
                    //Ben trai
                    switch (history.Rows[0].Field<string>("ClosingDate"))
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
                    switch (dt.Rows[0].Field<string>("ClosingDate"))
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

                if (dt.Rows[0].Field<string>("WorkType") != history.Rows[0].Field<string>("WorkType"))
                {
                    xlWorkSheet.Cells[57, "E"] = "☑就労形態";
                    xlWorkSheet.Cells[57, "E"].Font.Bold = true;
                    //Ben trai
                    switch (history.Rows[0].Field<string>("WorkType"))
                    {
                        case "請負":
                            xlWorkSheet.Cells[57, "P"] = "派遣　　・　　☑請負";
                            break;
                        case "派遣":
                            xlWorkSheet.Cells[57, "P"] = "☑派遣　　・　　請負";
                            break;
                    }
                    // Ben phai
                    switch (dt.Rows[0].Field<string>("WorkType"))
                    {
                        case "請負":
                            xlWorkSheet.Cells[57, "BB"] = "派遣　　・　　☑請負";
                            break;
                        case "派遣":
                            xlWorkSheet.Cells[57, "BB"] = "☑派遣　　・　　請負";
                            break;
                    }
                }
                string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if (zipcode != (history.Rows[0].Field<int?>("ZipCode")).ToString() || history.Rows[0].Field<string>("Address1") != dt.Rows[0].Field<string>("Address1") ||
                     history.Rows[0].Field<string>("Address2") != dt.Rows[0].Field<string>("Address2") || history.Rows[0].Field<string>("Address3") != dt.Rows[0].Field<string>("Address3")
                    || history.Rows[0].Field<string>("Address4") != dt.Rows[0].Field<string>("Address4") || history.Rows[0].Field<string>("Address5") != dt.Rows[0].Field<string>("Address5"))
                {
                    //Bảng bên phải
                    xlWorkSheet.Cells[33, "E"] = "☑住民票住所";
                    xlWorkSheet.Cells[33, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[33, "BE"] = zipcode;
                    // xlWorkSheet.Cells[36, "BE"] = TB_ZipCode.Text;
                    string address2 = dt.Rows[0].Field<string>("Address1") + dt.Rows[0].Field<string>("Address2") +
                   dt.Rows[0].Field<string>("Address3") + dt.Rows[0].Field<string>("Address5") + dt.Rows[0].Field<string>("Address5");
                    xlWorkSheet.Cells[33, "BK"] = address2;
                    //// Bang ben trái
                    xlWorkSheet.Cells[33, "S"] = (history.Rows[0].Field<int?>("ZipCode")).ToString();
                    xlWorkSheet.Cells[36, "S"] = (history.Rows[0].Field<int?>("ZipCode")).ToString();
                    string address1 = history.Rows[0].Field<string>("Address1") + history.Rows[0].Field<string>("Address2") +
                   history.Rows[0].Field<string>("Address3") + history.Rows[0].Field<string>("Address5") + history.Rows[0].Field<string>("Address5");
                    xlWorkSheet.Cells[33, "Y"] = address1;
                }

                if (dt.Rows[0].Field<string>("TravelType") != history.Rows[0].Field<string>("TravelType"))
                {   ///Bảng bên phải
                    xlWorkSheet.Cells[39, "E"] = "☑通勤/入寮";
                    xlWorkSheet.Cells[39, "E"].Font.Bold = true;
                    switch (dt.Rows[0].Field<string>("TravelType"))
                    {
                        case "入寮":
                            xlWorkSheet.Cells[39, "BK"] = "☑";
                            break;
                        case "通勤":
                            xlWorkSheet.Cells[39, "BB"] = "☑";
                            break;
                    }
                    // Bảng bên trái
                    switch (history.Rows[0].Field<string>("TravelType"))
                    {
                        case "入寮":
                            xlWorkSheet.Cells[39, "Y"] = "☑";
                            break;
                        case "通勤":
                            xlWorkSheet.Cells[39, "P"] = "☑";
                            break;
                    }
                }

                if (dt.Rows[0].Field<string>("EmployTime1") != history.Rows[0].Field<string>("EmployTime1") || dt.Rows[0].Field<string>("EmployTime2") != history.Rows[0].Field<string>("EmployTime2"))
                {   //Bảng bên phải
                    string temp_time1 = dt.Rows[0].Field<string>("EmployTime1");
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[42, "BB"] = a + b;
                    xlWorkSheet.Cells[42, "E"] = "☑雇用期間";
                    xlWorkSheet.Cells[42, "E"].Font.Bold = true;

                    //  Bảng bên trái
                    if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
                    {
                        string temp_time11 = history.Rows[0].Field<string>("EmployTime1");
                        string[] Time11_temps = temp_time11.Split('/');
                        string a1 = "平成 " + (Convert.ToInt32(Time11_temps[0]) - 1988).ToString() + " 年 " + Time11_temps[1] + " 月 " + Time11_temps[2] + " 日から ";

                        string temp_time21 = history.Rows[0].Field<string>("EmployTime2");
                        string[] Time21_temps = temp_time21.Split('/');
                        string b1 = "平成 " + (Convert.ToInt32(Time21_temps[0]) - 1988).ToString() + " 年 " + Time21_temps[1] + " 月 " + Time21_temps[2] + " 日まで";

                        xlWorkSheet.Cells[42, "P"] = a1 + b1;
                    }
                }

                if (dt.Rows[0].Field<string>("CardTime") != history.Rows[0].Field<string>("CardTime") || dt.Rows[0].Field<string>("CardTimeOut") != history.Rows[0].Field<string>("CardTimeOut"))
                {   // Bên phải
                    string temp_time1 = dt.Rows[0].Field<string>("CardTime");
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = dt.Rows[0].Field<string>("CardTimeOut");
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[48, "BB"] = a + b;
                    xlWorkSheet.Cells[48, "E"] = "☑在留期間";
                    xlWorkSheet.Cells[48, "E"].Font.Bold = true;

                    //Bên trái
                    if (history.Rows[0].Field<string>("Nationality") != string.Empty)
                    {
                        string temp_time11 = history.Rows[0].Field<string>("CardTime");
                        string[] Time11_temps = temp_time11.Split('/');
                        string a1 = "平成 " + (Convert.ToInt32(Time11_temps[0]) - 1988).ToString() + " 年 " + Time11_temps[1] + " 月 " + Time11_temps[2] + " 日から ";

                        string temp_time21 = history.Rows[0].Field<string>("CardTimeOut");
                        string[] Time21_temps = temp_time21.Split('/');
                        string b1 = "平成 " + (Convert.ToInt32(Time21_temps[0]) - 1988).ToString() + " 年 " + Time21_temps[1] + " 月 " + Time21_temps[2] + " 日まで";

                        xlWorkSheet.Cells[48, "P"] = a1 + b1;
                    }
                }

                if (history.Rows[0].Field<string>("CompanyName") != dt.Rows[0].Field<string>("CompanyName"))
                {   //Bên phải
                    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[51, "E"] = "☑企業名";
                    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[51, "BB"] = dt.Rows[0].Field<string>("CompanyName");
                    //Bên trái
                    xlWorkSheet.Cells[51, "P"] = history.Rows[0].Field<string>("CompanyName");
                }

                if (history.Rows[0].Field<string>("ShiharaiType") != dt.Rows[0].Field<string>("ShiharaiType"))
                {
                    xlWorkSheet.Cells[60, "BF"] = dt.Rows[0].Field<string>("ShiharaiType");
                    xlWorkSheet.Cells[60, "E"] = "☑賃金支払形態";
                    xlWorkSheet.Cells[60, "E"].Font.Bold = true;
                    //Bên trái
                    xlWorkSheet.Cells[60, "T"] = history.Rows[0].Field<string>("ShiharaiType");
                }

                if (history.Rows[0].Field<string>("Tax") != dt.Rows[0].Field<string>("Tax"))
                {   //Bên phải
                    xlWorkSheet.Cells[60, "CA"] = dt.Rows[0].Field<string>("Tax");
                    xlWorkSheet.Cells[62, "E"] = "☑税適";
                    xlWorkSheet.Cells[62, "E"].Font.Bold = true;
                    // Bên trái
                    xlWorkSheet.Cells[60, "AM"] = history.Rows[0].Field<string>("Tax");
                }

                if (history.Rows[0].Field<int?>("HakenRyokin").ToString() != dt.Rows[0].Field<int?>("HakenRyokin").ToString() || history.Rows[0].Field<string>("HakenRyokinType").ToString() != dt.Rows[0].Field<string>("HakenRyokinType").ToString())
                {   //Bên phải
                    xlWorkSheet.Cells[67, "BF"] = dt.Rows[0].Field<int?>("HakenRyokin");
                    xlWorkSheet.Cells[67, "BT"] = dt.Rows[0].Field<string>("HakenRyokinType");
                    xlWorkSheet.Cells[67, "E"] = "☑派遣・請金";
                    xlWorkSheet.Cells[67, "E"].Font.Bold = true;
                    //Bên trái
                    xlWorkSheet.Cells[67, "T"] = history.Rows[0].Field<int?>("HakenRyokin");
                    xlWorkSheet.Cells[67, "AX"] = history.Rows[0].Field<string>("HakenRyokinType");
                }
                if (history.Rows[0].Field<int?>("DormitoryFee").ToString() != dt.Rows[0].Field<int?>("DormitoryFee").ToString())
                {
                    xlWorkSheet.Cells[85, "AJ"] = history.Rows[0].Field<int?>("DormitoryFee");
                    xlWorkSheet.Cells[85, "R"] = "寮費";

                    xlWorkSheet.Cells[85, "BO"] = dt.Rows[0].Field<int?>("DormitoryFee").ToString();
                    xlWorkSheet.Cells[85, "BD"] = "寮費";
                    xlWorkSheet.Cells[85, "E"] = "☑給与控除額";
                    xlWorkSheet.Cells[85, "E"].Font.Bold = true;
                }

                if (history.Rows[0].Field<int?>("Chingin").ToString() != dt.Rows[0].Field<int?>("Chingin").ToString() || history.Rows[0].Field<string>("ChinginType").ToString() != dt.Rows[0].Field<string>("ChinginType").ToString())
                {   //Bên phải
                    xlWorkSheet.Cells[70, "BF"] = dt.Rows[0].Field<int?>("Chingin");
                    xlWorkSheet.Cells[70, "BT"] = dt.Rows[0].Field<string>("ChinginType");
                    xlWorkSheet.Cells[70, "E"] = "☑賃金";
                    xlWorkSheet.Cells[70, "E"].Font.Bold = true;
                    //Bên trái
                    xlWorkSheet.Cells[70, "T"] = history.Rows[0].Field<int?>("Chingin");
                    xlWorkSheet.Cells[70, "AX"] = history.Rows[0].Field<string>("ChinginType");
                }

                if (old_tsukinTeate != TB_TsukinTeate.Text)
                {   // Bên phải
                    xlWorkSheet.Cells[76, "BB"] = TB_TsukinTeate.Text;
                    xlWorkSheet.Cells[76, "E"] = "☑通勤手当";
                    xlWorkSheet.Cells[76, "E"].Font.Bold = true;
                    // bên trái
                    xlWorkSheet.Cells[76, "P"] = old_tsukinTeate;
                }

                if (history.Rows[0].Field<string>("BankCode") != dt.Rows[0].Field<string>("BankCode") || history.Rows[0].Field<string>("BranchCode") != dt.Rows[0].Field<string>("BranchCode")
                    || history.Rows[0].Field<string>("BankName") != dt.Rows[0].Field<string>("BankName") || history.Rows[0].Field<string>("BankNameType") != dt.Rows[0].Field<string>("BankNameType") ||
                    history.Rows[0].Field<string>("BranchName") != dt.Rows[0].Field<string>("BranchName") || history.Rows[0].Field<string>("BranchNameType") != dt.Rows[0].Field<string>("BranchNameType"))
                {   // Bên phải
                    xlWorkSheet.Cells[91, "E"] = "☑銀行/支店名";
                    xlWorkSheet.Cells[91, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[91, "BH"] = dt.Rows[0].Field<string>("BankCode");
                    xlWorkSheet.Cells[91, "CC"] = dt.Rows[0].Field<string>("BranchCode");
                    xlWorkSheet.Cells[92, "BB"] = dt.Rows[0].Field<string>("BankName");
                    xlWorkSheet.Cells[92, "BT"] = dt.Rows[0].Field<string>("BankNameType");
                    xlWorkSheet.Cells[92, "BW"] = dt.Rows[0].Field<string>("BranchName");
                    xlWorkSheet.Cells[92, "CO"] = dt.Rows[0].Field<string>("BranchNameType");
                    // Bên trái(cũ)
                    xlWorkSheet.Cells[91, "U"] = history.Rows[0].Field<string>("BankCode");
                    xlWorkSheet.Cells[91, "AN"] = history.Rows[0].Field<string>("BranchCode");
                    xlWorkSheet.Cells[92, "P"] = history.Rows[0].Field<string>("BankName");
                    xlWorkSheet.Cells[92, "AF"] = history.Rows[0].Field<string>("BankNameType");
                    xlWorkSheet.Cells[92, "AI"] = history.Rows[0].Field<string>("BranchName");
                    xlWorkSheet.Cells[92, "AY"] = history.Rows[0].Field<string>("BranchNameType");

                }
                if (history.Rows[0].Field<string>("AccountName") != dt.Rows[0].Field<string>("AccountName"))
                {
                    xlWorkSheet.Cells[95, "E"] = "☑口座名義（カナ";
                    xlWorkSheet.Cells[95, "E"].Font.Bold = true;
                    //Ben phai
                    xlWorkSheet.Cells[95, "BB"] = dt.Rows[0].Field<string>("AccountName");
                    //Ben trai
                    xlWorkSheet.Cells[95, "P"] = history.Rows[0].Field<string>("AccountName");
                }
                if (history.Rows[0].Field<string>("AccountCode1") != dt.Rows[0].Field<string>("AccountCode1") || history.Rows[0].Field<string>("AccountCode2") != dt.Rows[0].Field<string>("AccountCode2") ||
                history.Rows[0].Field<string>("AccountCode3") != dt.Rows[0].Field<string>("AccountCode3") || history.Rows[0].Field<string>("AccountCode4") != dt.Rows[0].Field<string>("AccountCode4") ||
                history.Rows[0].Field<string>("AccountCode5") != dt.Rows[0].Field<string>("AccountCode5") || history.Rows[0].Field<string>("AccountCode6") != dt.Rows[0].Field<string>("AccountCode6") ||
                history.Rows[0].Field<string>("AccountCode7") != dt.Rows[0].Field<string>("AccountCode7") || history.Rows[0].Field<string>("AccountCode8") != dt.Rows[0].Field<string>("AccountCode8"))
                {
                    xlWorkSheet.Cells[98, "E"] = "☑口座番号";
                    xlWorkSheet.Cells[98, "E"].Font.Bold = true;
                    //Bên phải
                    xlWorkSheet.Cells[98, "BF"] = dt.Rows[0].Field<string>("AccountCode1") + dt.Rows[0].Field<string>("AccountCode2") +
                   dt.Rows[0].Field<string>("AccountCode3") + dt.Rows[0].Field<string>("AccountCode4") + dt.Rows[0].Field<string>("AccountCode5")
                   + dt.Rows[0].Field<string>("AccountCode6") + dt.Rows[0].Field<string>("AccountCode7") + dt.Rows[0].Field<string>("AccountCode8");
                    //Bên trái
                    xlWorkSheet.Cells[98, "T"] = history.Rows[0].Field<string>("AccountCode1") + history.Rows[0].Field<string>("AccountCode2") +
                   history.Rows[0].Field<string>("AccountCode3") + history.Rows[0].Field<string>("AccountCode4") + history.Rows[0].Field<string>("AccountCode5")
                   + history.Rows[0].Field<string>("AccountCode6") + history.Rows[0].Field<string>("AccountCode7") + history.Rows[0].Field<string>("AccountCode8");
                }

                if (dt.Rows[0].Field<string>("Kouyouhoken") != " " && dt.Rows[0].Field<string>("Kouyouhoken") != history.Rows[0].Field<string>("Kouyouhoken"))
                {   //Bên phải
                    xlWorkSheet.Cells[101, "E"] = "☑雇用保険";
                    xlWorkSheet.Cells[101, "E"].Font.Bold = true;
                    string temp = dt.Rows[0].Field<string>("Kouyouhoken");
                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[101, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[101, "CJ"] = temps[1];
                    xlWorkSheet.Cells[101, "CN"] = temps[2];
                    //Bên trái
                    if (history.Rows[0].Field<string>("Kouyouhoken") != " ")
                    {
                        string temp1 = history.Rows[0].Field<string>("Kouyouhoken");

                        string[] temps1 = temp1.Split('/');
                        xlWorkSheet.Cells[101, "AJ"] = (Convert.ToInt32(temps1[0]) - 1988).ToString();
                        xlWorkSheet.Cells[101, "AO"] = temps1[1];
                        xlWorkSheet.Cells[101, "AT"] = temps1[2];
                    }
                }
                if (dt.Rows[0].Field<string>("Shakaihoken") != " " && dt.Rows[0].Field<string>("Shakaihoken") != history.Rows[0].Field<string>("Shakaihoken"))
                {   //Bên phải
                    xlWorkSheet.Cells[104, "E"] = "☑社会保険";
                    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                    string temp = dt.Rows[0].Field<string>("Shakaihoken");
                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[104, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[104, "CJ"] = temps[1];
                    xlWorkSheet.Cells[104, "CN"] = temps[2];
                    //Bên trái
                    if (history.Rows[0].Field<string>("Shakaihoken") != " ")
                    {
                        string temp1 = history.Rows[0].Field<string>("Shakaihoken");

                        string[] temps1 = temp1.Split('/');
                        xlWorkSheet.Cells[104, "AJ"] = (Convert.ToInt32(temps1[0]) - 1988).ToString();
                        xlWorkSheet.Cells[104, "AO"] = temps1[1];
                        xlWorkSheet.Cells[104, "AT"] = temps1[2];
                    }
                }

                if (TB_DependentPeople.Text != old_DependentPeople ||
                    TB_ResidentPeople.Text != old_ResidentPeople ||
                    TB_HealthInsurancePeople.Text != old_HealthInsurancePeople)
                {   //Bên phải
                    xlWorkSheet.Cells[104, "E"] = "☑扶 養 人 数";
                    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[107, "BK"] = TB_DependentPeople.Text;
                    xlWorkSheet.Cells[107, "BY"] = TB_ResidentPeople.Text;
                    xlWorkSheet.Cells[107, "CM"] = TB_HealthInsurancePeople.Text;
                    // Bên trái
                    xlWorkSheet.Cells[107, "X"] = old_DependentPeople;
                    xlWorkSheet.Cells[107, "AK"] = old_ResidentPeople;
                    xlWorkSheet.Cells[107, "AW"] = old_HealthInsurancePeople;
                }

                // Print out 1 copy to the default printer:
                xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                     Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                /////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////In ca 4 trang truoc nua///////////////////////

                //////////////////////////////IN TRANG 1-NYUSHANAIYOU////////////////////////////
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

                ///////////////////////////In TRANG 2-KEIYAKUSHO////////////////////////////////////////

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                string temp_ChangeDate1 = dt.Rows[0].Field<string>("ChangeDate");
                if (temp_ChangeDate1 != " ")
                {
                    string[] changedate = temp_ChangeDate1.Split('/');
                    xlWorkSheet.Cells[3, "AB"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[3, "AD"] = changedate[1];
                    xlWorkSheet.Cells[3, "AG"] = changedate[2];
                }
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
                string zipcode2 = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if (zipcode.Length == 7)
                {
                    string temp1_zipcode = zipcode2.Substring(0, 3);
                    string temp2_zipcode = zipcode2.Substring(3, 4);
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

                //////////////////////////IN TRANG 3-HOKEN////////////////////////////////////////////////////
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
                string birth1 = dt.Rows[0].Field<string>("Birth");
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
                string joindate1 = dt.Rows[0].Field<string>("InCompanyDate");
                if (joindate1 != " ")
                {
                    string[] joindate_temps = joindate1.Split('/');
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
                string zipcode1 = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if (zipcode.Length == 7)
                {
                    string temp1_zipcode = zipcode1.Substring(0, 3);
                    string temp2_zipcode = zipcode1.Substring(3, 4);
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

                ///////////////////////////////////IN TRANG4- KOUTSUHI/////////////////////////////////////////

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

                // Cleanup:
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);

                t.Abort(); //cho nay de huy thread
                MessageBox.Show("印刷準備完了");

                
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


        private void SplashScreen()
        {
            Application.Run(new GUI_SplashScreen());
        }

        private void Export() {
            Thread t = new Thread(new ThreadStart(SplashScreen));
            t.Start();
            Thread.Sleep(5000);

            BLL_Print bll_print = new BLL_Print();
            BLL_HandleFunc bll_handle = new BLL_HandleFunc();
            DataTable dt = bll_edit.EditForm(name);
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            try
            {

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(path + @"\File\template_export.xls",
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);

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

                string temp_ChangeDate = dt.Rows[0].Field<string>("ChangeDate");
                if (temp_ChangeDate != " ")
                {
                    string[] changedate = temp_ChangeDate.Split('/');
                    xlWorkSheet.Cells[21, "W"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[21, "AF"] = changedate[1];
                    xlWorkSheet.Cells[21, "AP"] = changedate[2];

                    xlWorkSheet.Cells[64, "BF"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[64, "BL"] = changedate[1];
                    xlWorkSheet.Cells[64, "BQ"] = changedate[2];
                }
                string temp_ChangeDateFrom = dt.Rows[0].Field<string>("ChangeDateFrom");
                if (temp_ChangeDateFrom != " ")
                {
                    string[] changedateFrom = temp_ChangeDateFrom.Split('/');
                    xlWorkSheet.Cells[24, "W"] = (Convert.ToInt32(changedateFrom[0]) - 1988).ToString();
                    xlWorkSheet.Cells[24, "AF"] = changedateFrom[1];
                    xlWorkSheet.Cells[24, "AP"] = dt.Rows[0].Field<string>("ClosingDate");

                }
                /////////////In cả hai bảng những chỗ thay đổi///////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////////////////

                foreach (DataRow row in dt.Rows)
                {
                    if (CB_TeateType.SelectedIndex > -1|| old_SeikinTeate != dt.Rows[0].Field<int>("SeikinTeate") ||
                        old_GaikinTeate!= dt.Rows[0].Field<int>("GaikinTeate") || old_YakushokuTeate != dt.Rows[0].Field<int>("GijutsuTeate") ||
                         old_ShikakuTeate != dt.Rows[0].Field<int>("ShikakuTeate") || old_YakushokuTeate != dt.Rows[0].Field<int>("YakushokuTeate") ||
                        old_EigyoTeate != dt.Rows[0].Field<int>("EigyoTeate") || old_KazokuTeate != dt.Rows[0].Field<int>("KazokuTeate") ||
                        old_JutakuTeate != dt.Rows[0].Field<int>("JutakuTeate") || old_BekkyoTeate!= dt.Rows[0].Field<int>("BekkyoTeate"))
                    {
                        switch (this.CB_TeateType.SelectedItem.ToString())
                        {
                            case "精勤手当":
                                // Ben trai
                                xlWorkSheet.Cells[79, "R"] = "精勤手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_SeikinTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;

                                // Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "精勤手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("SeikinTeate").ToString();
                                break;
                            case "外勤手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "外勤手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_GaikinTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                // Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "外勤手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("GaikinTeate").ToString();
                                break;
                            case "技術手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "技術手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_GijutsuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                // ben phai
                                xlWorkSheet.Cells[79, "BD"] = "技術手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("GijutsuTeate").ToString();
                                break;
                            case "資格手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "資格手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_ShikakuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "資格手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("ShikakuTeate").ToString();
                                break;
                            case "役職手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "役職手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_YakushokuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "役職手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("YakushokuTeate").ToString();
                                break;
                            case "営業・職務手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "営業・職務手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_EigyoTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "営業・職務手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("EigyoTeate").ToString();
                                break;
                            case "家族手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "家族手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_KazokuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "家族手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("KazokuTeate").ToString();
                                break;
                            case "住宅手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "住宅手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_JutakuTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "住宅手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("JutakuTeate").ToString();
                                break;
                            case "別居手当":
                                //Ben trai
                                xlWorkSheet.Cells[79, "R"] = "別居手当";
                                xlWorkSheet.Cells[79, "AJ"] = old_BekkyoTeate;
                                xlWorkSheet.Cells[79, "E"] = "☑手当額①";
                                xlWorkSheet.Cells[79, "E"].Font.Bold = true;
                                //Ben phai
                                xlWorkSheet.Cells[79, "BD"] = "別居手当";
                                xlWorkSheet.Cells[79, "BO"] = dt.Rows[0].Field<int>("BekkyoTeate").ToString();
                                break;
                        }
                    }
                }
                if (history.Rows[0].Field<string>("RomajiName") != dt.Rows[0].Field<string>("RomajiName"))
                {
                    xlWorkSheet.Cells[30, "P"] = history.Rows[0].Field<string>("RomajiName");
                    xlWorkSheet.Cells[30, "BB"] = dt.Rows[0].Field<string>("RomajiName");
                    xlWorkSheet.Cells[30, "E"] = "☑氏　名";
                    xlWorkSheet.Cells[30, "E"].Font.Bold = true;
                }
                if (dt.Rows[0].Field<string>("CardType") != history.Rows[0].Field<string>("CardType"))
                {
                    xlWorkSheet.Cells[45, "E"] = "☑在留資格";
                    xlWorkSheet.Cells[45, "E"].Font.Bold = true;
                    //Ben trai
                    switch (history.Rows[0].Field<string>("CardType"))
                    {
                        case "定住者":
                            xlWorkSheet.Cells[45, "P"] = "☑定・永・特永・日配・永配・その他（";
                            break;
                        case "永住者":
                            xlWorkSheet.Cells[45, "P"] = "定・☑永・特永・日配・永配・その他（";
                            break;
                        case "特別永住":
                            xlWorkSheet.Cells[45, "P"] = "定・永・☑特永・日配・永配・その他（";
                            break;
                        case "日本人配":
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・☑日配・永配・その他（";
                            break;
                        case "永住配":
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・日配・☑永配・その他（";
                            break;
                        default:
                            xlWorkSheet.Cells[45, "P"] = "定・永・特永・日配・永配・☑その他（";
                            xlWorkSheet.Cells[45, "AQ"] = dt.Rows[0].Field<string>("CardType");
                            break;
                    }
                    //Ben phai
                    switch (dt.Rows[0].Field<string>("CardType"))
                    {
                        case "定住者":
                            xlWorkSheet.Cells[45, "BB"] = "☑定・永・特永・日配・永配・その他（";
                            break;
                        case "永住者":
                            xlWorkSheet.Cells[45, "BB"] = "定・☑永・特永・日配・永配・その他（";
                            break;
                        case "特別永住":
                            xlWorkSheet.Cells[45, "BB"] = "定・永・☑特永・日配・永配・その他（";
                            break;
                        case "日本人配":
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・☑日配・永配・その他（";
                            break;
                        case "永住配":
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・日配・☑永配・その他（";
                            break;
                        default:
                            xlWorkSheet.Cells[45, "BB"] = "定・永・特永・日配・永配・☑その他（";
                            xlWorkSheet.Cells[45, "CD"] = CB_CardType.SelectedItem.ToString();
                            break;
                    }

                }
                if (dt.Rows[0].Field<string>("ClosingDate") != history.Rows[0].Field<string>("ClosingDate"))
                {
                    xlWorkSheet.Cells[54, "E"] = "☑締日";
                    xlWorkSheet.Cells[54, "E"].Font.Bold = true;
                    //Ben trai
                    switch (history.Rows[0].Field<string>("ClosingDate"))
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
                    switch (dt.Rows[0].Field<string>("ClosingDate"))
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

                if (dt.Rows[0].Field<string>("WorkType") != history.Rows[0].Field<string>("WorkType"))
                {
                    xlWorkSheet.Cells[57, "E"] = "☑就労形態";
                    xlWorkSheet.Cells[57, "E"].Font.Bold = true;
                    //Ben trai
                    switch (history.Rows[0].Field<string>("WorkType"))
                    {
                        case "請負":
                            xlWorkSheet.Cells[57, "P"] = "派遣　　・　　☑請負";
                            break;
                        case "派遣":
                            xlWorkSheet.Cells[57, "P"] = "☑派遣　　・　　請負";
                            break;
                    }
                    // Ben phai
                    switch (dt.Rows[0].Field<string>("WorkType"))
                    {
                        case "請負":
                            xlWorkSheet.Cells[57, "BB"] = "派遣　　・　　☑請負";
                            break;
                        case "派遣":
                            xlWorkSheet.Cells[57, "BB"] = "☑派遣　　・　　請負";
                            break;
                    }
                }
                string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if (zipcode != (history.Rows[0].Field<int?>("ZipCode")).ToString() || history.Rows[0].Field<string>("Address1") != dt.Rows[0].Field<string>("Address1") ||
                     history.Rows[0].Field<string>("Address2") != dt.Rows[0].Field<string>("Address2") || history.Rows[0].Field<string>("Address3") != dt.Rows[0].Field<string>("Address3")
                    || history.Rows[0].Field<string>("Address4") != dt.Rows[0].Field<string>("Address4") || history.Rows[0].Field<string>("Address5") != dt.Rows[0].Field<string>("Address5"))
                {
                    //Bảng bên phải
                    xlWorkSheet.Cells[33, "E"] = "☑住民票住所";
                    xlWorkSheet.Cells[33, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[33, "BE"] = zipcode;
                    // xlWorkSheet.Cells[36, "BE"] = TB_ZipCode.Text;
                    string address2 = dt.Rows[0].Field<string>("Address1") + dt.Rows[0].Field<string>("Address2") +
                   dt.Rows[0].Field<string>("Address3") + dt.Rows[0].Field<string>("Address5") + dt.Rows[0].Field<string>("Address5");
                    xlWorkSheet.Cells[33, "BK"] = address2;
                    //// Bang ben trái
                    xlWorkSheet.Cells[33, "S"] = (history.Rows[0].Field<int?>("ZipCode")).ToString();
                    xlWorkSheet.Cells[36, "S"] = (history.Rows[0].Field<int?>("ZipCode")).ToString();
                    string address1 = history.Rows[0].Field<string>("Address1") + history.Rows[0].Field<string>("Address2") +
                   history.Rows[0].Field<string>("Address3") + history.Rows[0].Field<string>("Address5") + history.Rows[0].Field<string>("Address5");
                    xlWorkSheet.Cells[33, "Y"] = address1;
                }

                if (dt.Rows[0].Field<string>("TravelType") != history.Rows[0].Field<string>("TravelType"))
                {   ///Bảng bên phải
                    xlWorkSheet.Cells[39, "E"] = "☑通勤/入寮";
                    xlWorkSheet.Cells[39, "E"].Font.Bold = true;
                    switch (dt.Rows[0].Field<string>("TravelType"))
                    {
                        case "入寮":
                            xlWorkSheet.Cells[39, "BK"] = "☑";
                            break;
                        case "通勤":
                            xlWorkSheet.Cells[39, "BB"] = "☑";
                            break;
                    }
                    // Bảng bên trái
                    switch (history.Rows[0].Field<string>("TravelType"))
                    {
                        case "入寮":
                            xlWorkSheet.Cells[39, "Y"] = "☑";
                            break;
                        case "通勤":
                            xlWorkSheet.Cells[39, "P"] = "☑";
                            break;
                    }
                }

                if (dt.Rows[0].Field<string>("EmployTime1") != history.Rows[0].Field<string>("EmployTime1") || dt.Rows[0].Field<string>("EmployTime2") != history.Rows[0].Field<string>("EmployTime2"))
                {   //Bảng bên phải
                    string temp_time1 = dt.Rows[0].Field<string>("EmployTime1");
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[42, "BB"] = a + b;
                    xlWorkSheet.Cells[42, "E"] = "☑雇用期間";
                    xlWorkSheet.Cells[42, "E"].Font.Bold = true;

                    //  Bảng bên trái
                    if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
                    {
                        string temp_time11 = history.Rows[0].Field<string>("EmployTime1");
                        string[] Time11_temps = temp_time11.Split('/');
                        string a1 = "平成 " + (Convert.ToInt32(Time11_temps[0]) - 1988).ToString() + " 年 " + Time11_temps[1] + " 月 " + Time11_temps[2] + " 日から ";

                        string temp_time21 = history.Rows[0].Field<string>("EmployTime2");
                        string[] Time21_temps = temp_time21.Split('/');
                        string b1 = "平成 " + (Convert.ToInt32(Time21_temps[0]) - 1988).ToString() + " 年 " + Time21_temps[1] + " 月 " + Time21_temps[2] + " 日まで";

                        xlWorkSheet.Cells[42, "P"] = a1 + b1;
                    }
                }

                if (dt.Rows[0].Field<string>("CardTime") != history.Rows[0].Field<string>("CardTime") || dt.Rows[0].Field<string>("CardTimeOut") != history.Rows[0].Field<string>("CardTimeOut"))
                {   // Bên phải
                    string temp_time1 = dt.Rows[0].Field<string>("CardTime");
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = dt.Rows[0].Field<string>("CardTimeOut");
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[48, "BB"] = a + b;
                    xlWorkSheet.Cells[48, "E"] = "☑在留期間";
                    xlWorkSheet.Cells[48, "E"].Font.Bold = true;

                    //Bên trái
                    if (history.Rows[0].Field<string>("Nationality") != string.Empty)
                    {
                        string temp_time11 = history.Rows[0].Field<string>("CardTime");
                        string[] Time11_temps = temp_time11.Split('/');
                        string a1 = "平成 " + (Convert.ToInt32(Time11_temps[0]) - 1988).ToString() + " 年 " + Time11_temps[1] + " 月 " + Time11_temps[2] + " 日から ";

                        string temp_time21 = history.Rows[0].Field<string>("CardTimeOut");
                        string[] Time21_temps = temp_time21.Split('/');
                        string b1 = "平成 " + (Convert.ToInt32(Time21_temps[0]) - 1988).ToString() + " 年 " + Time21_temps[1] + " 月 " + Time21_temps[2] + " 日まで";

                        xlWorkSheet.Cells[48, "P"] = a1 + b1;
                    }
                }

                if (history.Rows[0].Field<string>("CompanyName") != dt.Rows[0].Field<string>("CompanyName"))
                {   //Bên phải
                    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[51, "E"] = "☑企業名";
                    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[51, "BB"] = dt.Rows[0].Field<string>("CompanyName");
                    //Bên trái
                    xlWorkSheet.Cells[51, "P"] = history.Rows[0].Field<string>("CompanyName");
                }

                if (history.Rows[0].Field<string>("ShiharaiType") != dt.Rows[0].Field<string>("ShiharaiType"))
                {
                    xlWorkSheet.Cells[60, "BF"] = dt.Rows[0].Field<string>("ShiharaiType");
                    xlWorkSheet.Cells[60, "E"] = "☑賃金支払形態";
                    xlWorkSheet.Cells[60, "E"].Font.Bold = true;
                    //Bên trái
                    xlWorkSheet.Cells[60, "T"] = history.Rows[0].Field<string>("ShiharaiType");
                }

                if (history.Rows[0].Field<string>("Tax") != dt.Rows[0].Field<string>("Tax"))
                {   //Bên phải
                    xlWorkSheet.Cells[60, "CA"] = dt.Rows[0].Field<string>("Tax");
                    xlWorkSheet.Cells[62, "E"] = "☑税適";
                    xlWorkSheet.Cells[62, "E"].Font.Bold = true;
                    // Bên trái
                    xlWorkSheet.Cells[60, "AM"] = history.Rows[0].Field<string>("Tax");
                }

                if (history.Rows[0].Field<int?>("HakenRyokin").ToString() != dt.Rows[0].Field<int?>("HakenRyokin").ToString() || history.Rows[0].Field<string>("HakenRyokinType").ToString() != dt.Rows[0].Field<string>("HakenRyokinType").ToString())
                {   //Bên phải
                    xlWorkSheet.Cells[67, "BF"] = dt.Rows[0].Field<int?>("HakenRyokin");
                    xlWorkSheet.Cells[67, "BT"] = dt.Rows[0].Field<string>("HakenRyokinType");
                    xlWorkSheet.Cells[67, "E"] = "☑派遣・請金";
                    xlWorkSheet.Cells[67, "E"].Font.Bold = true;
                    //Bên trái
                    xlWorkSheet.Cells[67, "T"] = history.Rows[0].Field<int?>("HakenRyokin");
                    xlWorkSheet.Cells[67, "AX"] = history.Rows[0].Field<string>("HakenRyokinType");
                }
                if (history.Rows[0].Field<int?>("DormitoryFee").ToString() != dt.Rows[0].Field<int?>("DormitoryFee").ToString())
                {
                    xlWorkSheet.Cells[85, "AJ"] = history.Rows[0].Field<int?>("DormitoryFee");
                    xlWorkSheet.Cells[85, "R"] = "寮費";

                    xlWorkSheet.Cells[85, "BO"] = dt.Rows[0].Field<int?>("DormitoryFee").ToString();
                    xlWorkSheet.Cells[85, "BD"] = "寮費";
                    xlWorkSheet.Cells[85, "E"] = "☑給与控除額";
                    xlWorkSheet.Cells[85, "E"].Font.Bold = true;
                }

                if (history.Rows[0].Field<int?>("Chingin").ToString() != dt.Rows[0].Field<int?>("Chingin").ToString() || history.Rows[0].Field<string>("ChinginType").ToString() != dt.Rows[0].Field<string>("ChinginType").ToString())
                {   //Bên phải
                    xlWorkSheet.Cells[70, "BF"] = dt.Rows[0].Field<int?>("Chingin");
                    xlWorkSheet.Cells[70, "BT"] = dt.Rows[0].Field<string>("ChinginType");
                    xlWorkSheet.Cells[70, "E"] = "☑賃金";
                    xlWorkSheet.Cells[70, "E"].Font.Bold = true;
                    //Bên trái
                    xlWorkSheet.Cells[70, "T"] = history.Rows[0].Field<int?>("Chingin");
                    xlWorkSheet.Cells[70, "AX"] = history.Rows[0].Field<string>("ChinginType");
                }

                if (old_tsukinTeate != TB_TsukinTeate.Text)
                {   // Bên phải
                    xlWorkSheet.Cells[76, "BB"] = TB_TsukinTeate.Text;
                    xlWorkSheet.Cells[76, "E"] = "☑通勤手当";
                    xlWorkSheet.Cells[76, "E"].Font.Bold = true;
                    // bên trái
                    xlWorkSheet.Cells[76, "P"] = old_tsukinTeate;
                }

                if (history.Rows[0].Field<string>("BankCode") != dt.Rows[0].Field<string>("BankCode") || history.Rows[0].Field<string>("BranchCode") != dt.Rows[0].Field<string>("BranchCode")
                    || history.Rows[0].Field<string>("BankName") != dt.Rows[0].Field<string>("BankName") || history.Rows[0].Field<string>("BankNameType") != dt.Rows[0].Field<string>("BankNameType") ||
                    history.Rows[0].Field<string>("BranchName") != dt.Rows[0].Field<string>("BranchName") || history.Rows[0].Field<string>("BranchNameType") != dt.Rows[0].Field<string>("BranchNameType"))
                {   // Bên phải
                    xlWorkSheet.Cells[91, "E"] = "☑銀行/支店名";
                    xlWorkSheet.Cells[91, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[91, "BH"] = dt.Rows[0].Field<string>("BankCode");
                    xlWorkSheet.Cells[91, "CC"] = dt.Rows[0].Field<string>("BranchCode");
                    xlWorkSheet.Cells[92, "BB"] = dt.Rows[0].Field<string>("BankName");
                    xlWorkSheet.Cells[92, "BT"] = dt.Rows[0].Field<string>("BankNameType");
                    xlWorkSheet.Cells[92, "BW"] = dt.Rows[0].Field<string>("BranchName");
                    xlWorkSheet.Cells[92, "CO"] = dt.Rows[0].Field<string>("BranchNameType");
                    // Bên trái(cũ)
                    xlWorkSheet.Cells[91, "U"] = history.Rows[0].Field<string>("BankCode");
                    xlWorkSheet.Cells[91, "AN"] = history.Rows[0].Field<string>("BranchCode");
                    xlWorkSheet.Cells[92, "P"] = history.Rows[0].Field<string>("BankName");
                    xlWorkSheet.Cells[92, "AF"] = history.Rows[0].Field<string>("BankNameType");
                    xlWorkSheet.Cells[92, "AI"] = history.Rows[0].Field<string>("BranchName");
                    xlWorkSheet.Cells[92, "AY"] = history.Rows[0].Field<string>("BranchNameType");

                }
                if (history.Rows[0].Field<string>("AccountName") != dt.Rows[0].Field<string>("AccountName"))
                {
                    xlWorkSheet.Cells[95, "E"] = "☑口座名義（カナ";
                    xlWorkSheet.Cells[95, "E"].Font.Bold = true;
                    //Ben phai
                    xlWorkSheet.Cells[95, "BB"] = dt.Rows[0].Field<string>("AccountName");
                    //Ben trai
                    xlWorkSheet.Cells[95, "P"] = history.Rows[0].Field<string>("AccountName");
                }
                if (history.Rows[0].Field<string>("AccountCode1") != dt.Rows[0].Field<string>("AccountCode1") || history.Rows[0].Field<string>("AccountCode2") != dt.Rows[0].Field<string>("AccountCode2") ||
                history.Rows[0].Field<string>("AccountCode3") != dt.Rows[0].Field<string>("AccountCode3") || history.Rows[0].Field<string>("AccountCode4") != dt.Rows[0].Field<string>("AccountCode4") ||
                history.Rows[0].Field<string>("AccountCode5") != dt.Rows[0].Field<string>("AccountCode5") || history.Rows[0].Field<string>("AccountCode6") != dt.Rows[0].Field<string>("AccountCode6") ||
                history.Rows[0].Field<string>("AccountCode7") != dt.Rows[0].Field<string>("AccountCode7") || history.Rows[0].Field<string>("AccountCode8") != dt.Rows[0].Field<string>("AccountCode8"))
                {
                    xlWorkSheet.Cells[98, "E"] = "☑口座番号";
                    xlWorkSheet.Cells[98, "E"].Font.Bold = true;
                    //Bên phải
                    xlWorkSheet.Cells[98, "BF"] = dt.Rows[0].Field<string>("AccountCode1") + dt.Rows[0].Field<string>("AccountCode2") +
                   dt.Rows[0].Field<string>("AccountCode3") + dt.Rows[0].Field<string>("AccountCode4") + dt.Rows[0].Field<string>("AccountCode5")
                   + dt.Rows[0].Field<string>("AccountCode6") + dt.Rows[0].Field<string>("AccountCode7") + dt.Rows[0].Field<string>("AccountCode8");
                    //Bên trái
                    xlWorkSheet.Cells[98, "T"] = history.Rows[0].Field<string>("AccountCode1") + history.Rows[0].Field<string>("AccountCode2") +
                   history.Rows[0].Field<string>("AccountCode3") + history.Rows[0].Field<string>("AccountCode4") + history.Rows[0].Field<string>("AccountCode5")
                   + history.Rows[0].Field<string>("AccountCode6") + history.Rows[0].Field<string>("AccountCode7") + history.Rows[0].Field<string>("AccountCode8");
                }

                if (dt.Rows[0].Field<string>("Kouyouhoken") != " " && dt.Rows[0].Field<string>("Kouyouhoken") != history.Rows[0].Field<string>("Kouyouhoken"))
                {   //Bên phải
                    xlWorkSheet.Cells[101, "E"] = "☑雇用保険";
                    xlWorkSheet.Cells[101, "E"].Font.Bold = true;
                    string temp = dt.Rows[0].Field<string>("Kouyouhoken");
                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[101, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[101, "CJ"] = temps[1];
                    xlWorkSheet.Cells[101, "CN"] = temps[2];
                    //Bên trái
                    if (history.Rows[0].Field<string>("Kouyouhoken") != " ")
                    {
                        string temp1 = history.Rows[0].Field<string>("Kouyouhoken");

                        string[] temps1 = temp1.Split('/');
                        xlWorkSheet.Cells[101, "AJ"] = (Convert.ToInt32(temps1[0]) - 1988).ToString();
                        xlWorkSheet.Cells[101, "AO"] = temps1[1];
                        xlWorkSheet.Cells[101, "AT"] = temps1[2];
                    }
                }
                if (dt.Rows[0].Field<string>("Shakaihoken") != " " && dt.Rows[0].Field<string>("Shakaihoken") != history.Rows[0].Field<string>("Shakaihoken"))
                {   //Bên phải
                    xlWorkSheet.Cells[104, "E"] = "☑社会保険";
                    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                    string temp = dt.Rows[0].Field<string>("Shakaihoken");
                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[104, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[104, "CJ"] = temps[1];
                    xlWorkSheet.Cells[104, "CN"] = temps[2];
                    //Bên trái
                    if (history.Rows[0].Field<string>("Shakaihoken") != " ")
                    {
                        string temp1 = history.Rows[0].Field<string>("Shakaihoken");

                        string[] temps1 = temp1.Split('/');
                        xlWorkSheet.Cells[104, "AJ"] = (Convert.ToInt32(temps1[0]) - 1988).ToString();
                        xlWorkSheet.Cells[104, "AO"] = temps1[1];
                        xlWorkSheet.Cells[104, "AT"] = temps1[2];
                    }
                }

                if (TB_DependentPeople.Text != old_DependentPeople ||
                    TB_ResidentPeople.Text != old_ResidentPeople ||
                    TB_HealthInsurancePeople.Text != old_HealthInsurancePeople)
                {   //Bên phải
                    xlWorkSheet.Cells[104, "E"] = "☑扶 養 人 数";
                    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[107, "BK"] = TB_DependentPeople.Text;
                    xlWorkSheet.Cells[107, "BY"] = TB_ResidentPeople.Text;
                    xlWorkSheet.Cells[107, "CM"] = TB_HealthInsurancePeople.Text;
                    // Bên trái
                    xlWorkSheet.Cells[107, "X"] = old_DependentPeople;
                    xlWorkSheet.Cells[107, "AK"] = old_ResidentPeople;
                    xlWorkSheet.Cells[107, "AW"] = old_HealthInsurancePeople;
                }


                /////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////In ca 4 trang truoc nua///////////////////////

                //////////////////////////////IN TRANG 1-NYUSHANAIYOU////////////////////////////
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

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


                ///////////////////////////In TRANG 2-KEIYAKUSHO////////////////////////////////////////

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);

                string temp_ChangeDate1 = dt.Rows[0].Field<string>("ChangeDate");
                if (temp_ChangeDate1 != " ")
                {
                    string[] changedate = temp_ChangeDate1.Split('/');
                    xlWorkSheet.Cells[3, "AB"] = (Convert.ToInt32(changedate[0]) - 1988).ToString();
                    xlWorkSheet.Cells[3, "AD"] = changedate[1];
                    xlWorkSheet.Cells[3, "AG"] = changedate[2];
                }
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
                string zipcode2 = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if (zipcode.Length == 7)
                {
                    string temp1_zipcode = zipcode2.Substring(0, 3);
                    string temp2_zipcode = zipcode2.Substring(3, 4);
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

                //////////////////////////IN TRANG 3-HOKEN////////////////////////////////////////////////////
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);

                xlWorkSheet.Cells[2, "AP"] = dt.Rows[0].Field<string>("Position");
                xlWorkSheet.Cells[4, "AP"] = dt.Rows[0].Field<string>("CreatePeople");
                xlWorkSheet.Cells[11, "D"] = dt.Rows[0].Field<string>("IDCode");
                xlWorkSheet.Cells[12, "S"] = dt.Rows[0].Field<string>("RomajiName");
                xlWorkSheet.Cells[11, "S"] = dt.Rows[0].Field<string>("FuriganaName");
                xlWorkSheet.Cells[12, "AP"] = dt.Rows[0].Field<string>("Sex");
                xlWorkSheet.Cells[14, "S"] = dt.Rows[0].Field<string>("CompanyName");
                xlWorkSheet.Cells[15, "AL"] = dt.Rows[0].Field<string>("ClosingDate");
                //Birthday
                string birth1 = dt.Rows[0].Field<string>("Birth");
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
                string joindate1 = dt.Rows[0].Field<string>("InCompanyDate");
                if (joindate1 != " ")
                {
                    string[] joindate_temps = joindate1.Split('/');
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
                string zipcode1 = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if (zipcode.Length == 7)
                {
                    string temp1_zipcode = zipcode1.Substring(0, 3);
                    string temp2_zipcode = zipcode1.Substring(3, 4);
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


                ///////////////////////////////////IN TRANG4- KOUTSUHI/////////////////////////////////////////

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);

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

                t.Abort();
                ////////// show promt to save file
                System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
                saveDlg.InitialDirectory = @"C:\";
                saveDlg.Filter = "Excel files (*.xls)|*.xls";
                saveDlg.FilterIndex = 0;
                saveDlg.RestoreDirectory = true;
                saveDlg.Title = "Export Excel File To";
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
            catch (Exception eSavePrint)
            {
                // Cleanup Memory
                xlWorkBook.Close(0);
                xlApp.Quit();
                MessageBox.Show(eSavePrint.Message, "エラー！印刷できません！");
            }
        }
    }
}
