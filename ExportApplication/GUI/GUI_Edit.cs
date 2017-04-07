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
        // Enable and disable nhung truong duoc chon de edit
        public void TakeThis(IList<int> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                switch (list[i])
                {
                    case 0:
                        TB_IDCode.Enabled = true;
                        label1.Font = new Font(label1.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label1.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 1:
                        TB_RomajiName.Enabled = true;
                        label2.Font = new Font(label2.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label2.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 2:
                        TB_FuriganaName.Enabled = true;
                        label3.Font = new Font(label3.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label3.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 3:
                        TB_CompanyCode.Enabled = true;
                        label4.Font = new Font(label4.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label4.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 4:
                        TB_CompanyName.Enabled = true;
                        label5.Font = new Font(label5.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label5.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 5:
                        CB_Sex.Enabled = true;
                        label6.Font = new Font(label6.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label6.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 6:
                        CB_ShiharaiType.Enabled = true;
                        label7.Font = new Font(label7.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label7.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 7:
                        CB_ClosingDate.Enabled = true;
                        label8.Font = new Font(label8.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label8.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 8:
                        DTP_Birth.Enabled = true;
                        label9.Font = new Font(label9.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label9.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 9:
                        TB_Reason.Enabled = true;
                        label10.Font = new Font(label10.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label10.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 10:
                        DTP_ChangeDate.Enabled = true;
                        label11.Font = new Font(label11.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label11.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 11:
                        DTP_ChangeDateFrom.Enabled = true;
                        label12.Font = new Font(label12.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label12.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 12:
                        TB_ZipCode.Enabled = true;
                        TB_Address1.Enabled = true;
                        TB_Address3.Enabled = true;
                        TB_Address5.Enabled = true;
                        CB_Address2.Enabled = true;
                        CB_Address4.Enabled = true;
                        label24.Font = new Font(label24.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label24.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 13:
                        CB_TravelType.Enabled = true;
                        label22.Font = new Font(label22.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label22.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 14:
                        TB_KyuyoKojoGaku.Enabled = true;
                        label36.Font = new Font(label36.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label36.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 15:
                        TB_HakenRyokin.Enabled = true;
                        CB_HakenRyokinType.Enabled = true;
                        label27.Font = new Font(label27.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label27.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 16:
                        TB_TeateGaku.Enabled = true;
                        CB_TeateType.Enabled = true;
                        label34.Font = new Font(label34.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label34.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 17:
                        TB_TsukinTeate.Enabled = true;
                        label32.Font = new Font(label32.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label32.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 18:
                        TB_Chingin.Enabled = true;
                        CB_ChinginType.Enabled = true;
                        label30.Font = new Font(label30.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label30.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 19:
                        CB_WorkType.Enabled = true;
                        label15.Font = new Font(label15.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label15.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 20:
                        CB_Tax.Enabled = true;
                        label26.Font = new Font(label26.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label26.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 21:
                        CB_CardType.Enabled = true;
                        label20.Font = new Font(label20.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label20.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 22:
                        DTP_CardTimeStart.Enabled = true;
                        DTP_CardTimeOver.Enabled = true;
                        label14.Font = new Font(label14.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label14.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 23:
                        DTP_EmployTime1.Enabled = true;
                        DTP_EmployTime2.Enabled = true;
                        label21.Font = new Font(label21.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label21.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 24:
                        TB_WorkTime.Enabled = true;
                        label38.Font = new Font(label38.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label38.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 25:
                        TB_BankCode.Enabled = true;
                        label39.Font = new Font(label39.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label39.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 26:
                        TB_BranchCode.Enabled = true;
                        label40.Font = new Font(label40.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label40.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 27:
                        TB_BankName.Enabled = true;
                        CB_BankNameType.Enabled = true;
                        label41.Font = new Font(label41.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label41.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 28:
                        TB_BranchName.Enabled = true;
                        CB_BranchNameType.Enabled = true;
                        label42.Font = new Font(label42.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label42.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 29:
                        TB_AccountName.Enabled = true;
                        label43.Font = new Font(label43.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label43.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 30:
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
                    case 31:
                        DTP_KoyoHokenDate.Enabled = true;
                        label45.Font = new Font(label45.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label45.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 32:
                        DTP_CompanyInsureDate.Enabled = true;
                        label46.Font = new Font(label46.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label46.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 33:
                        TB_DependentPeople.Enabled = true;
                        //TB_ResidentPeople.Enabled = true;
                        //TB_HealthInsurancePeople.Enabled = true;
                        label47.Font = new Font(label47.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label47.ForeColor = System.Drawing.Color.Red;
                        break;

                }
            }
        }

        public GUI_Edit()
        {
            InitializeComponent();
        }
        // get name lay tu dataGridView de gui len DAL
        public static string name;
        public void funData(string text)
        {
            name = text;

        }
        //Load du thong tin cua nguoi duoc chon de edit
        public delegate void delPassData(string text);
        string old_tsukinTeate = string.Empty;
        string old_DependentPeople = string.Empty;
        string old_ResidentPeople = string.Empty;
        string old_HealthInsurancePeople = string.Empty;
        private void GUI_Edit_Load(object sender, EventArgs e)
        {
            DataTable dt = bll_edit.EditForm(name);
            old_tsukinTeate = dt.Rows[0].Field<int?>("TotalMoneyTrans").ToString();
            old_DependentPeople = dt.Rows[0].Field<int?>("DependentPeople").ToString();
            old_ResidentPeople = dt.Rows[0].Field<int?>("ResidentPeople").ToString();
            old_HealthInsurancePeople = dt.Rows[0].Field<int?>("HealthInsurancePeople").ToString();
            // MessageBox.Show(dt.Rows[0].Field<string>("RomajiName"));
            foreach (DataRow row in dt.Rows)
            {
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

        // Save button click-> save database to information DB

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
            string _TeateGaku = TB_TeateGaku.Text;
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

            DTO_Edit dto_edit = new DTO_Edit(_RomajiName, _IDCode, _FuriganaName, _CompanyName, _CompanyCode, _Sex, _ShiharaiType, _Tax, _Birth, _Reason,
            _ChangeDate, _ChangeDateFrom, _ZipCode, _Address1, _Address2, _Address3, _Address4, _Address5, _TravelType, _EmployTime1, _EmployTime2, _CardType, _CardTimeOver,
            _CardTimeStart, _WorkType, _ClosingDate, _HakenRyokin, _ChinginType, _HakenRyokinType, _Chingin, _TsukinTeate, _TeateType, _Genkaritsu,
            _TeateGaku, _KyuyoKojoGaku, _WorkTime, _BankName, _BankNameType, _BranchName, _BranchNameType, _AccountName, _BankCode,
            _BranchCode, _AccountCode, _AccountCode1, _AccountCode2, _AccountCode3, _AccountCode4, _AccountCode5, _AccountCode6, _AccountCode7,
            _CompanyInsureDate, _KoyoHokenDate, _DependentPeople, _ResidentPeople, _HealthInsurancePeople);
            return dto_edit;
        }

        // click save button
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

        // envent hiển thị ngày tháng như được chọn
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

        private void DTP_KoyoHokenDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_KoyoHokenDate.Format = DateTimePickerFormat.Long;
        }

        private void DTP_CompanyInsureDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_CompanyInsureDate.Format = DateTimePickerFormat.Long;
        }
        private void TB_TsukinTeate_Click(object sender, EventArgs e)
        {
            GUI_Travel gui_travel = new GUI_Travel(name);
            gui_travel.Show();
        }

        // Bắt sự kiện chỉ cho nhập vào số chứ k cho nhập chữ vào
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

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void panel3_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel3.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel2.ClientRectangle;
            rect.Width--;
            rect.Height--;
            Pen p = new Pen(Color.FromArgb(17, 168, 171), 1);
            e.Graphics.DrawRectangle(p, rect);
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

        private void Print() {
            DataTable dt = bll_edit.EditForm(name);
            // MessageBox.Show(dt.Rows[0].Field<string>("RomajiName"));

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
                }

                //Bang ben trai 
                xlWorkSheet.Cells[30, "P"] = dt.Rows[0].Field<string>("RomajiName");
                string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                xlWorkSheet.Cells[33, "S"] = zipcode;
                xlWorkSheet.Cells[36, "S"] = zipcode;

                //Address
                string address = dt.Rows[0].Field<string>("Address1") + dt.Rows[0].Field<string>("Address2") +
                    dt.Rows[0].Field<string>("Address3") + dt.Rows[0].Field<string>("Address4") + dt.Rows[0].Field<string>("Address5");
                xlWorkSheet.Cells[33, "Y"] = address;
                // xlWorkSheet.Cells[36, "Y"] = address;

                if (dt.Rows[0].Field<string>("TravelType") == "入寮")
                {
                    xlWorkSheet.Cells[39, "Y"] = "☑";
                }
                else
                {
                    xlWorkSheet.Cells[39, "P"] = "☑";
                }

                if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
                {
                    string temp_time1 = dt.Rows[0].Field<string>("EmployTime1");
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[42, "P"] = a + b;
                }

                if (dt.Rows[0].Field<string>("Nationality") != string.Empty)
                {
                    string temp_time1 = dt.Rows[0].Field<string>("CardTime");
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = dt.Rows[0].Field<string>("CardTimeOut");
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[48, "P"] = a + b;
                }

                xlWorkSheet.Cells[51, "P"] = dt.Rows[0].Field<string>("CompanyName");

                xlWorkSheet.Cells[60, "T"] = dt.Rows[0].Field<string>("ShiharaiType");
                xlWorkSheet.Cells[60, "AM"] = dt.Rows[0].Field<string>("Tax");

                xlWorkSheet.Cells[67, "T"] = dt.Rows[0].Field<int?>("HakenRyokin");
                xlWorkSheet.Cells[67, "AX"] = dt.Rows[0].Field<string>("HakenRyokinType");

                xlWorkSheet.Cells[70, "T"] = dt.Rows[0].Field<int?>("Chingin");
                xlWorkSheet.Cells[70, "AX"] = dt.Rows[0].Field<string>("ChinginType");

                xlWorkSheet.Cells[76, "P"] = old_tsukinTeate;

                xlWorkSheet.Cells[91, "U"] = dt.Rows[0].Field<string>("BankCode");
                xlWorkSheet.Cells[91, "AN"] = dt.Rows[0].Field<string>("BranchCode");
                xlWorkSheet.Cells[92, "P"] = dt.Rows[0].Field<string>("BankName");
                xlWorkSheet.Cells[92, "AF"] = dt.Rows[0].Field<string>("BankNameType");
                xlWorkSheet.Cells[92, "AI"] = dt.Rows[0].Field<string>("BranchName");
                xlWorkSheet.Cells[92, "AY"] = dt.Rows[0].Field<string>("BranchNameType");
                xlWorkSheet.Cells[95, "P"] = dt.Rows[0].Field<string>("AccountName");

                xlWorkSheet.Cells[98, "T"] = dt.Rows[0].Field<string>("AccountCode1") + dt.Rows[0].Field<string>("AccountCode2") +
                    dt.Rows[0].Field<string>("AccountCode3") + dt.Rows[0].Field<string>("AccountCode4") + dt.Rows[0].Field<string>("AccountCode5")
                    + dt.Rows[0].Field<string>("AccountCode6") + dt.Rows[0].Field<string>("AccountCode7") + dt.Rows[0].Field<string>("AccountCode8");

                if (dt.Rows[0].Field<string>("Kouyouhoken") != " ")
                {
                    string temp = dt.Rows[0].Field<string>("Kouyouhoken");

                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[101, "AJ"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[101, "AO"] = temps[1];
                    xlWorkSheet.Cells[101, "AT"] = temps[2];
                }
                if (dt.Rows[0].Field<string>("Shakaihoken") != " ")
                {
                    string temp = dt.Rows[0].Field<string>("Shakaihoken");

                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[104, "AJ"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[104, "AO"] = temps[1];
                    xlWorkSheet.Cells[104, "AT"] = temps[2];
                }

                xlWorkSheet.Cells[107, "X"] = old_DependentPeople;
                xlWorkSheet.Cells[107, "AK"] = old_ResidentPeople;
                xlWorkSheet.Cells[107, "AW"] = old_HealthInsurancePeople;


                ///////////////////////////////////////////
                if (TB_RomajiName.Text != dt.Rows[0].Field<string>("RomajiName"))
                {
                    xlWorkSheet.Cells[30, "BB"] = TB_RomajiName.Text;
                    xlWorkSheet.Cells[30, "E"] = "☑氏　名";
                    xlWorkSheet.Cells[30, "E"].Font.Bold = true;
                }

                if (zipcode != TB_ZipCode.Text || TB_Address1.Text != dt.Rows[0].Field<string>("Address1") ||
                    CB_Address2.SelectedItem.ToString() != dt.Rows[0].Field<string>("Address2") || TB_Address3.Text != dt.Rows[0].Field<string>("Address3")
                    || CB_Address4.SelectedItem.ToString() != dt.Rows[0].Field<string>("Address4") || TB_Address5.Text != dt.Rows[0].Field<string>("Address5"))
                {
                    xlWorkSheet.Cells[33, "E"] = "☑住民票住所";
                    xlWorkSheet.Cells[33, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[33, "BE"] = TB_ZipCode.Text;
                    // xlWorkSheet.Cells[36, "BE"] = TB_ZipCode.Text;
                    string address2 = TB_Address1.Text + CB_Address2.SelectedItem.ToString() +
                    TB_Address3.Text + CB_Address4.SelectedItem.ToString() + TB_Address5.Text;
                    xlWorkSheet.Cells[33, "BK"] = address2;
                }


                // xlWorkSheet.Cells[36, "BK"] = address2;


                if (dt.Rows[0].Field<string>("TravelType") != CB_TravelType.SelectedItem.ToString())
                {
                    xlWorkSheet.Cells[39, "E"] = "☑通勤/入寮";
                    xlWorkSheet.Cells[39, "E"].Font.Bold = true;
                    if (CB_TravelType.SelectedItem.ToString() == "入寮")
                    {
                        xlWorkSheet.Cells[39, "BK"] = "☑";
                    }
                    else
                    {
                        xlWorkSheet.Cells[39, "BB"] = "☑";
                    }

                }

                if (dt.Rows[0].Field<string>("EmployTime1") != CheckDateTime(DTP_EmployTime1))
                {
                    string temp_time1 = CheckDateTime(DTP_EmployTime1);
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = CheckDateTime(DTP_EmployTime2);
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[42, "BB"] = a + b;
                    xlWorkSheet.Cells[42, "E"] = "☑雇用期間";
                    xlWorkSheet.Cells[42, "E"].Font.Bold = true;
                }

                if (dt.Rows[0].Field<string>("CardTime") != CheckDateTime(DTP_CardTimeStart))
                {
                    string temp_time1 = CheckDateTime(DTP_CardTimeStart);
                    string[] Time1_temps = temp_time1.Split('/');
                    string a = "平成 " + (Convert.ToInt32(Time1_temps[0]) - 1988).ToString() + " 年 " + Time1_temps[1] + " 月 " + Time1_temps[2] + " 日から ";

                    string temp_time2 = CheckDateTime(DTP_CardTimeOver);
                    string[] Time2_temps = temp_time2.Split('/');
                    string b = "平成 " + (Convert.ToInt32(Time2_temps[0]) - 1988).ToString() + " 年 " + Time2_temps[1] + " 月 " + Time2_temps[2] + " 日まで";

                    xlWorkSheet.Cells[48, "BB"] = a + b;
                    xlWorkSheet.Cells[48, "E"] = "☑在留期間";
                    xlWorkSheet.Cells[48, "E"].Font.Bold = true;
                }

                if (TB_CompanyName.Text != dt.Rows[0].Field<string>("CompanyName"))
                {
                    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[51, "E"] = "☑企業名";
                    xlWorkSheet.Cells[51, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[51, "BB"] = TB_CompanyName.Text;
                }

                if (CB_ShiharaiType.SelectedItem.ToString() != dt.Rows[0].Field<string>("ShiharaiType"))
                {
                    xlWorkSheet.Cells[60, "BF"] = CB_ShiharaiType.SelectedItem.ToString();
                    xlWorkSheet.Cells[60, "E"] = "☑賃金支払形態";
                    xlWorkSheet.Cells[60, "E"].Font.Bold = true;
                }

                if (CB_Tax.SelectedItem.ToString() != dt.Rows[0].Field<string>("Tax"))
                {
                    xlWorkSheet.Cells[60, "CA"] = CB_Tax.SelectedItem.ToString();
                    xlWorkSheet.Cells[62, "E"] = "☑税適";
                    xlWorkSheet.Cells[62, "E"].Font.Bold = true;
                }

                if (TB_HakenRyokin.Text != dt.Rows[0].Field<int?>("HakenRyokin").ToString())
                {
                    xlWorkSheet.Cells[67, "BF"] = TB_HakenRyokin.Text;
                    xlWorkSheet.Cells[67, "BT"] = CB_HakenRyokinType.SelectedItem.ToString();
                    xlWorkSheet.Cells[67, "E"] = "☑派遣・請金";
                    xlWorkSheet.Cells[67, "E"].Font.Bold = true;
                }

                if (TB_Chingin.Text != dt.Rows[0].Field<int?>("Chingin").ToString())
                {
                    xlWorkSheet.Cells[70, "BF"] = TB_Chingin.Text;
                    xlWorkSheet.Cells[70, "BT"] = dt.Rows[0].Field<string>("ChinginType");
                    xlWorkSheet.Cells[70, "E"] = "☑賃金";
                    xlWorkSheet.Cells[70, "E"].Font.Bold = true;
                }

                if (old_tsukinTeate != TB_TsukinTeate.Text)
                {
                    xlWorkSheet.Cells[76, "BB"] = TB_TsukinTeate.Text;
                    xlWorkSheet.Cells[76, "E"] = "☑通勤手当";
                    xlWorkSheet.Cells[76, "E"].Font.Bold = true;
                }

                if (TB_BankCode.Text != dt.Rows[0].Field<string>("BankCode") || TB_BranchCode.Text != dt.Rows[0].Field<string>("BranchCode")
                    || TB_BankName.Text != dt.Rows[0].Field<string>("BankName") || CB_BankNameType.SelectedItem.ToString() != dt.Rows[0].Field<string>("BankNameType") ||
                    TB_BranchName.Text != dt.Rows[0].Field<string>("BranchName") || CB_BranchNameType.SelectedItem.ToString() != dt.Rows[0].Field<string>("BranchNameType")
                    || TB_AccountName.Text != dt.Rows[0].Field<string>("AccountName"))
                {
                    xlWorkSheet.Cells[91, "E"] = "☑銀行/支店名";
                    xlWorkSheet.Cells[91, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[91, "BH"] = TB_BankCode.Text;
                    xlWorkSheet.Cells[91, "CC"] = TB_BranchCode.Text;
                    xlWorkSheet.Cells[92, "BB"] = TB_BankName.Text;
                    xlWorkSheet.Cells[92, "BT"] = CB_BankNameType.SelectedItem.ToString();
                    xlWorkSheet.Cells[92, "BW"] = TB_BranchName.Text;
                    xlWorkSheet.Cells[92, "CO"] = CB_BranchNameType.SelectedItem.ToString();
                    xlWorkSheet.Cells[95, "BB"] = TB_AccountName.Text;
                    xlWorkSheet.Cells[98, "BF"] = TB_AccountCode.Text + TB_AccountCode1.Text + TB_AccountCode2.Text +
                    TB_AccountCode3.Text + TB_AccountCode4.Text + TB_AccountCode5.Text + TB_AccountCode6.Text + TB_AccountCode7.Text;
                }


                if (CheckDateTime(DTP_KoyoHokenDate) != " " && CheckDateTime(DTP_KoyoHokenDate) != dt.Rows[0].Field<string>("Kouyouhoken"))
                {
                    xlWorkSheet.Cells[101, "E"] = "☑雇用保険";
                    xlWorkSheet.Cells[101, "E"].Font.Bold = true;
                    string temp = CheckDateTime(DTP_KoyoHokenDate);
                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[101, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[101, "CJ"] = temps[1];
                    xlWorkSheet.Cells[101, "CN"] = temps[2];
                }
                if (CheckDateTime(DTP_CompanyInsureDate) != " " && CheckDateTime(DTP_KoyoHokenDate) != dt.Rows[0].Field<string>("Shakaihoken"))
                {
                    xlWorkSheet.Cells[104, "E"] = "☑社会保険";
                    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                    string temp = CheckDateTime(DTP_CompanyInsureDate);
                    string[] temps = temp.Split('/');
                    xlWorkSheet.Cells[104, "CF"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[104, "CJ"] = temps[1];
                    xlWorkSheet.Cells[104, "CN"] = temps[2];
                }

                if (TB_DependentPeople.Text != old_DependentPeople ||
                    TB_ResidentPeople.Text != old_ResidentPeople ||
                    TB_HealthInsurancePeople.Text != old_HealthInsurancePeople)
                {
                    xlWorkSheet.Cells[104, "E"] = "☑扶 養 人 数";
                    xlWorkSheet.Cells[104, "E"].Font.Bold = true;
                    xlWorkSheet.Cells[107, "BK"] = TB_DependentPeople.Text;
                    xlWorkSheet.Cells[107, "BY"] = TB_ResidentPeople.Text;
                    xlWorkSheet.Cells[107, "CM"] = TB_HealthInsurancePeople.Text;
                }



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

        private void TB_DependentPeople_Click(object sender, EventArgs e)
        {
            GUI_Dependent obj = new GUI_Dependent(name);
            obj.Show();
        }


    }
}
