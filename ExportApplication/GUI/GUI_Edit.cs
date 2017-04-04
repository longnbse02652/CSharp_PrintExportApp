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

namespace ExportApplication
{
    public partial class GUI_Edit : Form
    {
        BLL_Edit bll_edit = new BLL_Edit();
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
                        TB_Address.Enabled = true;
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
                        label47.Font = new Font(label47.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label47.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 34:
                        TB_ResidentPeople.Enabled = true;
                        label48.Font = new Font(label48.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label48.ForeColor = System.Drawing.Color.Red;
                        break;
                    case 35:
                        TB_HealthInsurancePeople.Enabled = true;
                        label49.Font = new Font(label49.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        label49.ForeColor = System.Drawing.Color.Red;
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
        private void GUI_Edit_Load(object sender, EventArgs e)
        {
            DataTable dt = bll_edit.EditForm(name);
           // MessageBox.Show(dt.Rows[0].Field<string>("RomajiName"));
            foreach (DataRow row in dt.Rows)
            {
                TB_RomajiName.Text = dt.Rows[0].Field<string>("RomajiName");
                TB_IDCode.Text = dt.Rows[0].Field<string>("IDCode");
                TB_FuriganaName.Text = dt.Rows[0].Field<string>("FuriganaName");
                TB_CompanyCode.Text = dt.Rows[0].Field<string>("CompanyCode");
                TB_CompanyName.Text = dt.Rows[0].Field<string>("CompanyName");
                TB_Reason.Text = dt.Rows[0].Field<string>("Reason");
                TB_Address.Text = dt.Rows[0].Field<string>("Address1");// la cho hang t2 trong bang edit
                TB_Address3.Text = dt.Rows[0].Field<string>("Address3");
             //   TB_Address1.Text = dt.Rows[0].Field<string>("Address1");
                TB_Address5.Text = dt.Rows[0].Field<string>("Address5");
                TB_TeateGaku.Text = dt.Rows[0].Field<string>("TeateGaku");
                TB_BankName.Text = dt.Rows[0].Field<string>("BankName");
                TB_BranchName.Text = dt.Rows[0].Field<string>("BranchName");
                TB_AccountName.Text = dt.Rows[0].Field<string>("AccountName");
                TB_BankCode.Text = dt.Rows[0].Field<string>("BankCode");
                TB_BranchCode.Text = dt.Rows[0].Field<string>("BranchCode");
                TB_AccountCode.Text = dt.Rows[0].Field<string>("AccountCode");
                CB_TravelType.Text = dt.Rows[0].Field<string>("TravelType");
                CB_BranchNameType.Text = dt.Rows[0].Field<string>("BranchNameType");
                CB_BankNameType.Text = dt.Rows[0].Field<string>("BankNameType");
              //  CB_TeateType.Text = dt.Rows[0].Field<string>("TeateType");
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
                object tsukinteate = row["TsukinTeate"];
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
                else { TB_TsukinTeate.Text = dt.Rows[0].Field<int>("TsukinTeate").ToString(); }

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
            string _Address = TB_Address.Text;
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
            string _CompanyInsureDate = DTP_CompanyInsureDate.Text;
            string _KoyoHokenDate = DTP_KoyoHokenDate.Text;
            string _Birth = DTP_Birth.Text;
            string _Reason = TB_Reason.Text;
            string _ChangeDate = DTP_ChangeDate.Text;
            string _ChangeDateFrom = DTP_ChangeDateFrom.Text;
            string _EmployTime1 = DTP_EmployTime1.Text;
            string _EmployTime2 = DTP_EmployTime2.Text;
            string _CardTimeOver = DTP_CardTimeOver.Text;
            string _CardTimeStart = DTP_CardTimeStart.Text;

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
            if (CB_Address2.SelectedIndex == -1) { _Address2 = ""; } else { _Address2 = CB_Address2.SelectedItem.ToString(); }
            if (CB_Address4.SelectedIndex == -1) { _Address4 = ""; } else { _Address4 = CB_Address4.SelectedItem.ToString(); }
            if (CB_WorkType.SelectedIndex == -1) { _WorkType = ""; } else { _WorkType = CB_WorkType.SelectedItem.ToString(); }
            if (CB_Tax.SelectedIndex == -1) { _Tax = ""; } else { _Tax = CB_Tax.SelectedItem.ToString(); }
            if (CB_ShiharaiType.SelectedIndex == -1) { _ShiharaiType = ""; } else { _ShiharaiType = CB_ShiharaiType.SelectedItem.ToString(); }
            if (CB_Sex.SelectedIndex == -1) { _Sex = ""; } else { _Sex = CB_Sex.SelectedItem.ToString(); }
            if (CB_TravelType.SelectedIndex == -1) { _TravelType = ""; } else { _TravelType = CB_TravelType.SelectedItem.ToString(); }
            if (CB_CardType.SelectedIndex == -1) { _CardType = ""; } else { _CardType = CB_CardType.SelectedItem.ToString(); }
            if (CB_ClosingDate.SelectedIndex == -1) { _ClosingDate = ""; } else { _ClosingDate = CB_ClosingDate.SelectedItem.ToString(); }
            if (CB_HakenRyokinType.SelectedIndex == -1) { _HakenRyokinType = ""; } else { _HakenRyokinType = CB_HakenRyokinType.SelectedItem.ToString(); }
            if (CB_ChinginType.SelectedIndex == -1) { _ChinginType = ""; } else { _ChinginType = CB_ChinginType.SelectedItem.ToString(); }
            if (CB_TeateType.SelectedIndex == -1) { _TeateType = ""; } else { _TeateType = CB_TeateType.SelectedItem.ToString(); }
            if (CB_BankNameType.SelectedIndex == -1) { _BankNameType = ""; } else { _BankNameType = CB_BankNameType.SelectedItem.ToString(); }
            if (CB_BranchNameType.SelectedIndex == -1) { _BranchNameType = ""; } else { _BranchNameType = CB_BranchNameType.SelectedItem.ToString(); }

            DTO_Edit dto_edit = new DTO_Edit(_RomajiName, _IDCode, _FuriganaName, _CompanyName, _CompanyCode, _Sex, _ShiharaiType, _Tax, _Birth, _Reason,
            _ChangeDate, _ChangeDateFrom, _ZipCode, _Address, _Address2, _Address3, _Address4, _Address5, _TravelType, _EmployTime1, _EmployTime2, _CardType, _CardTimeOver,
            _CardTimeStart, _WorkType, _ClosingDate, _HakenRyokin, _ChinginType, _HakenRyokinType, _Chingin, _TsukinTeate, _TeateType, _Genkaritsu,
            _TeateGaku, _KyuyoKojoGaku, _WorkTime, _BankName, _BankNameType, _BranchName, _BranchNameType, _AccountName, _BankCode,
            _BranchCode, _AccountCode, _CompanyInsureDate, _KoyoHokenDate, _DependentPeople, _ResidentPeople, _HealthInsurancePeople);
            return dto_edit;
        }

        // click save button
        private void btSave_Click_1(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("保存を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_edit.Insert(updateData()))
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

        // envent hiển thị ngày tháng như được chọn
        private void DTP_EmployTime1_ValueChanged(object sender, EventArgs e)
        {
            DTP_EmployTime1.Format = DateTimePickerFormat.Custom;
            DTP_EmployTime1.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_Birth_ValueChanged(object sender, EventArgs e)
        {
            DTP_Birth.Format = DateTimePickerFormat.Custom;
            DTP_Birth.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_ChangeDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_ChangeDate.Format = DateTimePickerFormat.Custom;
            DTP_ChangeDate.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_ChangeDateFrom_ValueChanged(object sender, EventArgs e)
        {
            DTP_ChangeDateFrom.Format = DateTimePickerFormat.Custom;
            DTP_ChangeDateFrom.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_EmployTime2_ValueChanged(object sender, EventArgs e)
        {
            DTP_EmployTime2.Format = DateTimePickerFormat.Custom;
            DTP_EmployTime2.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_CardTimeStart_ValueChanged(object sender, EventArgs e)
        {
            DTP_CardTimeStart.Format = DateTimePickerFormat.Custom;
            DTP_CardTimeStart.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_CardTimeOver_ValueChanged(object sender, EventArgs e)
        {
            DTP_CardTimeOver.Format = DateTimePickerFormat.Custom;
            DTP_CardTimeOver.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_KoyoHokenDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_KoyoHokenDate.Format = DateTimePickerFormat.Custom;
            DTP_KoyoHokenDate.CustomFormat = "yyyy年MM月dd日";
        }

        private void DTP_CompanyInsureDate_ValueChanged(object sender, EventArgs e)
        {
            DTP_CompanyInsureDate.Format = DateTimePickerFormat.Custom;
            DTP_CompanyInsureDate.CustomFormat = "yyyy年MM月dd日";
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

 
    }
}
