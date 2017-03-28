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
    public partial class GUI_AddNew : Form
    {
        BLL_AllInfor bll_allInfo = new BLL_AllInfor();
        BLL_HandleFunc bll_handleFunc = new BLL_HandleFunc(); //tạo object của class này ra để format trước khi chuyển xuống database
        DTO_AllInfor dto_allInfo;
        public GUI_AddNew()
        {
            InitializeComponent();
        }
        
        private void GUI_AddNew_Load(object sender, EventArgs e)
        {
            this.ActiveControl = tb_IDCode;
            dtp_Birth.CustomFormat = " ";
            dtp_Birth.Format = DateTimePickerFormat.Custom;
            dtp_InCompanyDate.CustomFormat = " ";
            dtp_InCompanyDate.Format = DateTimePickerFormat.Custom;
            dtp_CardTime.CustomFormat = " ";
            dtp_CardTime.Format = DateTimePickerFormat.Custom;
            dtp_CardTimeOut.CustomFormat = " ";
            dtp_CardTimeOut.Format = DateTimePickerFormat.Custom;
            dtp_EmployTime1.CustomFormat = " ";
            dtp_EmployTime1.Format = DateTimePickerFormat.Custom;
            dtp_EmployTime2.CustomFormat = " ";
            dtp_EmployTime2.Format = DateTimePickerFormat.Custom;
            dtp_InHouseDate.CustomFormat = " ";
            dtp_InHouseDate.Format = DateTimePickerFormat.Custom;
            dtp_kouyouhoken.CustomFormat = " ";
            dtp_kouyouhoken.Format = DateTimePickerFormat.Custom;
            dtp_shakaihoken.CustomFormat = " ";
            dtp_shakaihoken.Format = DateTimePickerFormat.Custom;
            dtp_DependentPeopleBirth1.CustomFormat = " ";
            dtp_DependentPeopleBirth1.Format = DateTimePickerFormat.Custom;
            dtp_DependentPeopleBirth2.CustomFormat = " ";
            dtp_DependentPeopleBirth2.Format = DateTimePickerFormat.Custom;
            dtp_DependentPeopleBirth3.CustomFormat = " ";
            dtp_DependentPeopleBirth3.Format = DateTimePickerFormat.Custom;
            dtp_DependentPeopleBirth4.CustomFormat = " ";
            dtp_DependentPeopleBirth4.Format = DateTimePickerFormat.Custom;
            dtp_DependentPeopleBirth5.CustomFormat = " ";
            dtp_DependentPeopleBirth5.Format = DateTimePickerFormat.Custom;
            dtp_DependentPeopleBirth6.CustomFormat = " ";
            dtp_DependentPeopleBirth6.Format = DateTimePickerFormat.Custom;
        }

        //nút này để lưu dữ liệu vào database
        private void bt_Save_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("登録を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_allInfo.Insert(getAllData()))
                {
                    MessageBox.Show("登録しました。");
                    GUI_Main obj = (GUI_Main)Application.OpenForms["GUI_Main"];
                    obj.LoadGridView();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("登録は失敗しました。");
                }
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //MessageBox.Show(bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtpicker_Birth) + tb_FuriganaName.Text + tb_RomajiName.Text);
                //cancel Save 
            }

        }

        public DTO_AllInfor getAllData() {
            if (tb_RomajiName.Text != string.Empty && tb_FuriganaName.Text != string.Empty)
            {
                string _idCode = tb_IDCode.Text;
                string _romaji = tb_RomajiName.Text;
                string _furigana = tb_FuriganaName.Text;
                string _sex = cb_Sex.SelectedItem.ToString();
                int _age;
                bool result_age = int.TryParse(tb_Age.Text, out _age);
                string _birth = CheckDateTime(dtp_Birth);
                string _nationality = tb_Nationality.Text;
                string _inCompanyDate = CheckDateTime(dtp_InCompanyDate);
                string _cardType = cb_CardType.SelectedItem.ToString();
                string _cardTime = CheckDateTime(dtp_CardTime);
                string _cardTimeOut = CheckDateTime(dtp_CardTimeOut);
                string _outTime = cb_OutTime.SelectedItem.ToString();
                string _companyCode = tb_CompanyCode.Text;
                string _companyName = tb_CompanyName.Text;
                string _workType = cb_WorkType.SelectedItem.ToString();
                string _closingDate = cb_ClosingDate.SelectedItem.ToString();
                int _zipCode;
                bool result_zipcode = int.TryParse(tb_ZipCode.Text, out _zipCode);
                string _address = tb_Address.Text;
                string _mobliePhone = tb_MobliePhone.Text;
                string _phone = tb_Phone.Text;
                string _createPeople = tb_CreatePeople.Text;
                string _position = tb_Position.Text;

                dto_allInfo = new DTO_AllInfor(_idCode, _romaji, _furigana, _sex, _age, _birth, _nationality,
                        _inCompanyDate, _cardType, _cardTime, _cardTimeOut, _outTime, _companyCode, _companyName, _workType,
                        _closingDate, _zipCode, _address, _mobliePhone, _phone, _createPeople, _position);

            }
            else {
                MessageBox.Show("名前を入力してください");
            }
            
            return dto_allInfo;
        }
        
        private void bt_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Validate ko nhap name, age, sex
        private void tb_RomajiName_Validating(object sender, CancelEventArgs e)
        {
            bll_handleFunc.ValidateControls(tb_RomajiName, errorProvider);
        }

        private void tb_FuriganaName_Validating(object sender, CancelEventArgs e)
        {
            bll_handleFunc.ValidateControls(tb_FuriganaName, errorProvider);
        }

        private void tb_Age_Validating(object sender, CancelEventArgs e)
        {
            bll_handleFunc.ValidateControls(tb_Age, errorProvider);
        }

        private void cb_Sex_Validating(object sender, CancelEventArgs e)
        {
            bll_handleFunc.ValidateControls(cb_Sex, errorProvider);
        }

        //Xu ly cho 基本賃金
        private void cb_SalaryType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cb_SalaryType.SelectedIndex) { 
                case 0:
                    lb_ChinginType.Text = "月";
                    break;
                case 1:
                    lb_ChinginType.Text = "日";
                    break;
                case 2:
                    lb_ChinginType.Text = "時";
                    break;
                case 3:
                    lb_ChinginType.Text = "";
                    break;
            }
        }

        //Xu ly cho 通勤形態
        private void cb_TravelType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_TravelType.SelectedIndex == 0)
            {
                tlp_TravelType.Enabled = false;
                tb_HouseName.Text = "";
                tb_Room.Text = "";
                dtp_InHouseDate.CustomFormat = " ";
                dtp_InHouseDate.Format = DateTimePickerFormat.Custom;
            }
            else {
                tlp_TravelType.Enabled = true;
            }
        
        }

        //xu ly datetimepicker phong truong hop user ko nhap 
        public string CheckDateTime(DateTimePicker dtp) { 
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

        private void dtp_Birth_ValueChanged(object sender, EventArgs e)
        {
            dtp_Birth.Format = DateTimePickerFormat.Long;
        }

        private void dtp_InCompanyDate_ValueChanged(object sender, EventArgs e)
        {
            dtp_InCompanyDate.Format = DateTimePickerFormat.Long;
        }

        private void dtp_CardTime_ValueChanged(object sender, EventArgs e)
        {
            dtp_CardTime.Format = DateTimePickerFormat.Long;
        }

        private void dtp_CardTimeOut_ValueChanged(object sender, EventArgs e)
        {
            dtp_CardTimeOut.Format = DateTimePickerFormat.Long;
        }

        private void dtp_EmployTime1_ValueChanged(object sender, EventArgs e)
        {
            dtp_EmployTime1.Format = DateTimePickerFormat.Long;
        }

        private void dtp_EmployTime2_ValueChanged(object sender, EventArgs e)
        {
            dtp_EmployTime2.Format = DateTimePickerFormat.Long;
        }

        private void dtp_InHouseDate_ValueChanged(object sender, EventArgs e)
        {
            dtp_InHouseDate.Format = DateTimePickerFormat.Long;
        }

        private void dtp_kouyouhoken_ValueChanged(object sender, EventArgs e)
        {
            dtp_kouyouhoken.Format = DateTimePickerFormat.Long;
        }

        private void dtp_shakaihoken_ValueChanged(object sender, EventArgs e)
        {
            dtp_shakaihoken.Format = DateTimePickerFormat.Long;
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

    }
}
