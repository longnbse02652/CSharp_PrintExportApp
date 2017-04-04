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
            dtp_CardTimeStart.CustomFormat = " ";
            dtp_CardTimeStart.Format = DateTimePickerFormat.Custom;
            dtp_CardTimeOver.CustomFormat = " ";
            dtp_CardTimeOver.Format = DateTimePickerFormat.Custom;
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
            
            if (tb_RomajiName.Text != string.Empty)
            {
                string _idCode = tb_IDCode.Text;
                string _romaji = tb_RomajiName.Text;
                string _furigana = tb_FuriganaName.Text;
                string _sex;
                if (cb_Sex.SelectedIndex != -1)
                {
                    _sex = cb_Sex.SelectedItem.ToString();
                }
                else
                {
                    _sex = string.Empty;
                }
                string _birth = CheckDateTime(dtp_Birth);
                string _nationality = tb_Nationality.Text;
                string _inCompanyDate = CheckDateTime(dtp_InCompanyDate);
                string _cardType;
                if (cb_CardType.SelectedIndex != -1)
                {
                    _cardType = cb_CardType.SelectedItem.ToString();
                }
                else
                {
                    _cardType = string.Empty;
                }
                string _cardTime = CheckDateTime(dtp_CardTimeStart);
                string _cardTimeOut = CheckDateTime(dtp_CardTimeOver);
                string _outTime = string.Empty;
                if (cb_OutTime.SelectedIndex != -1)
                {
                    _outTime = cb_OutTime.SelectedItem.ToString();
                }
                else
                {
                    _outTime = string.Empty;
                }
                string _companyCode = tb_CompanyCode.Text;
                string _companyName = tb_CompanyName.Text;
                string _workType;
                if (cb_WorkType.SelectedIndex != -1)
                {
                    _workType = cb_WorkType.SelectedItem.ToString();
                }
                else
                {
                    _workType = string.Empty;
                }
                string _closingDate;
                if (cb_ClosingDate.SelectedIndex != -1)
                {
                    _closingDate = cb_ClosingDate.SelectedItem.ToString();
                }
                else
                {
                    _closingDate = string.Empty;
                }
                int _zipCode;
                bool result_zipcode = int.TryParse(tb_ZipCode.Text, out _zipCode);
                string _address1 = tb_Address1.Text;
                string _address2;
                if (cb_Address2.SelectedIndex != -1 && cb_Address2.SelectedText != "")
                {
                    _address2 = cb_Address2.SelectedItem.ToString();
                }
                else
                {
                    _address2 = string.Empty;
                }
                string _address3 = tb_Address3.Text;
                string _address4;
                if (cb_Address4.SelectedIndex != -1 && cb_Address4.SelectedText != "")
                {
                    _address4 = cb_Address4.SelectedItem.ToString();
                }
                else
                {
                    _address4 = string.Empty;
                }
                string _address5 = tb_Address5.Text;
                string _mobliePhone = tb_MobliePhone.Text;
                string _phone = tb_Phone.Text;
                string _createPeople = tb_CreatePeople.Text;
                string _position;
                if (cb_Position.SelectedIndex != -1 && cb_Position.SelectedText !="")
                {
                    _position = cb_Position.SelectedItem.ToString();
                }
                else
                {
                    _position = string.Empty;
                }
                int _hakenRyokin;
                bool result_hakenRyokin = int.TryParse(tb_HakenRyokin.Text, out _hakenRyokin);
                string _hakenRyokinType = cb_HakenRyokinType.SelectedItem.ToString();
                string _shiharaiType = cb_ShiharaiType.SelectedItem.ToString();
                string _tax = cb_Tax.SelectedItem.ToString();
                string _salaryType = cb_SalaryType.SelectedItem.ToString();
                int _basicSalary;
                bool result_basicSalary = int.TryParse(tb_BasicSalary.Text, out _basicSalary);
                int _seikinTeate;
                bool result_seikinTeate = int.TryParse(tb_SeikinTeate.Text,out _seikinTeate);
                int _gaikinTeate;
                bool result_gaikinTeate = int.TryParse(tb_GaikinTeate.Text,out _gaikinTeate);
                int _gijutsuTeate;
                bool result_gijutsuTeate = int.TryParse(tb_GijutsuTeate.Text,out _gijutsuTeate);
                int _shikakuTeate;
                bool result_shikakuTeate = int.TryParse(tb_ShikakuTeate.Text,out _shikakuTeate);
                int _yakushokuTeate;
                bool result_yakushokuTeate = int.TryParse(tb_YakushokuTeate.Text,out _yakushokuTeate);
                int _eigyoTeate;
                bool result_eigyoTeate = int.TryParse(tb_EigyoTeate.Text,out _eigyoTeate);
                int _kazokuTeate;
                bool result_kazokuTeate = int.TryParse(tb_KazokuTeate.Text,out _kazokuTeate);
                int _jutakuTeate;
                bool result_jutakuTeate = int.TryParse(tb_JutakuTeate.Text,out _jutakuTeate);
                int _bekkyoTeate;
                bool result_bekkyoTeate = int.TryParse(tb_BekkyoTeate.Text,out _bekkyoTeate);
                int _tsukinTeate;
                bool result_tsukinTeate = int.TryParse(tb_TsukinTeate.Text,out _tsukinTeate);
                int _park;
                bool result_park = int.TryParse(tb_Park.Text,out _park);
                int _dormitoryFee;
                bool result_dormitoryFee = int.TryParse(tb_DormitoryFee.Text,out _dormitoryFee);
                int _waterFee;
                bool result_waterFee = int.TryParse(tb_WaterFee.Text,out _waterFee);
                string _employStatus = cb_EmployStatus.SelectedItem.ToString() ;
                string _employTime1 = CheckDateTime(dtp_EmployTime1);
                string _employTime2 = CheckDateTime(dtp_EmployTime2);
                string _bankName = tb_BankName.Text;
                string _bankNameType = cb_BankNameType.SelectedItem.ToString();
                string _branchName = tb_BranchName.Text;
                string _branchNameType = cb_BranchNameType.SelectedItem.ToString();
                string _accountName = tb_AccountName.Text;
                string _bankCode = tb_BankCode.Text;
                string _branchCode = tb_BranchCode.Text;
                string _accountCode1 = tb_AccountCode1.Text;
                string _accountCode2 = tb_AccountCode2.Text;
                string _accountCode3 = tb_AccountCode3.Text;
                string _accountCode4 = tb_AccountCode4.Text;
                string _accountCode5 = tb_AccountCode5.Text;
                string _accountCode6 = tb_AccountCode6.Text;
                string _accountCode7 = tb_AccountCode7.Text;
                string _accountCode8 = tb_AccountCode8.Text;
                string _travelType = cb_TravelType.SelectedItem.ToString();
                string _houseName = tb_HouseName.Text;
                string _room = tb_Room.Text;
                string _inHouseDate = CheckDateTime(dtp_InHouseDate);
                string _kouyouhokenDate = CheckDateTime(dtp_kouyouhoken);
                string _shakaihokenDate = CheckDateTime(dtp_shakaihoken);
                int _dependentPeople;
                bool result_dependentPeople = int.TryParse(tb_DependentPeople.Text,out _dependentPeople);
                int _residentPeople;
                bool result_residentPeople = int.TryParse(tb_ResidentPeople.Text,out _residentPeople);
                int _healthInsurancePeople;
                bool result_healthInsurancePeople = int.TryParse(tb_HealthInsurancePeople.Text,out _healthInsurancePeople);

                string _contractType = string.Empty;
                if (cb_ContractType.SelectedIndex != -1)
                {
                    _contractType = cb_ContractType.SelectedItem.ToString();
                }
                else
                {
                    _contractType = string.Empty;
                }
                string _contractRequire = string.Empty;
                if (cb_ContractRequire.SelectedIndex != -1)
                {
                    _contractRequire = cb_ContractRequire.SelectedItem.ToString();
                }
                else
                {
                    _contractRequire = string.Empty;
                }
                string _myCompany = tb_MyCompany.Text;
                string _workContent = tb_WorkContent.Text;
                string _workTime1 = tb_WorkTime1.Text;
                string _workTime2 = tb_WorkTime2.Text;
                string _workTime3 = tb_WorkTime3.Text;
                string _workTime4 = tb_WorkTime4.Text;
                string _relaxTime = tb_RelaxTime.Text;
                string _insureCard = string.Empty;
                if (cb_InsureCard.SelectedIndex != -1)
                {
                    _insureCard = cb_InsureCard.SelectedItem.ToString();
                }
                else
                {
                    _insureCard = string.Empty;
                }
                string _pastCompany1 = tb_PastCompany1.Text;
                string _nienhieu1 = cb_Nienhieu1.SelectedItem.ToString();
                int _beginYear1;
                bool result_beginYear1 = int.TryParse(tb_BeginYear1.Text,out _beginYear1);
                int _beginMonth1;
                bool result_beginMonth1 = int.TryParse(tb_BeginMonth1.Text,out _beginMonth1);
                int _endYear1;
                bool result_endYear1 = int.TryParse(tb_EndYear1.Text,out _endYear1);
                int _endMonth1;
                bool result_endMonth1 = int.TryParse(tb_EndMonth1.Text,out _endMonth1);
                string _pastCompany2 = tb_PastCompany2.Text;
                string _nienhieu2 = cb_Nienhieu2.SelectedItem.ToString();
                int _beginYear2;
                bool result_beginYear2 = int.TryParse(tb_BeginYear2.Text,out _beginYear2);
                int _beginMonth2;
                bool result_beginMonth2 = int.TryParse(tb_BeginMonth2.Text,out _beginMonth2);
                int _endYear2;
                bool result_endYear2 = int.TryParse(tb_EndYear2.Text,out _endYear2);
                int _endMonth2;
                bool result_endMonth2 = int.TryParse(tb_EndMonth2.Text,out _endMonth2);
                string _pensionBook = string.Empty;
                if (cb_PensionBook.SelectedIndex != -1)
                {
                    _pensionBook = cb_PensionBook.SelectedItem.ToString();
                }
                else
                {
                    _pensionBook = string.Empty;
                }
                string _dependentPeopleKana1 = tb_DependentPeopleKana1.Text;
                string _dependentPeopleShimei1 = tb_DependentPeopleShimei1.Text;
                string _dependentPeopleBirth1 = CheckDateTime(dtp_DependentPeopleBirth1);
                string _relationship1 = tb_Relationship1.Text;
                string _living1 = cb_Living1.SelectedItem.ToString();
                string _dependentPeopleKana2 = tb_DependentPeopleKana2.Text;
                string _dependentPeopleShimei2 = tb_DependentPeopleShimei2.Text;
                string _dependentPeopleBirth2 = CheckDateTime(dtp_DependentPeopleBirth2);
                string _relationship2 = tb_Relationship2.Text;
                string _living2 = cb_Living2.SelectedItem.ToString();
                string _dependentPeopleKana3 = tb_DependentPeopleKana3.Text;
                string _dependentPeopleShimei3 = tb_DependentPeopleShimei3.Text;
                string _dependentPeopleBirth3 = CheckDateTime(dtp_DependentPeopleBirth3);
                string _relationship3 = tb_Relationship3.Text;
                string _living3 = cb_Living3.SelectedItem.ToString();
                string _dependentPeopleKana4 = tb_DependentPeopleKana4.Text;
                string _dependentPeopleShimei4 = tb_DependentPeopleShimei4.Text;
                string _dependentPeopleBirth4 = CheckDateTime(dtp_DependentPeopleBirth4);
                string _relationship4 = tb_Relationship4.Text;
                string _living4 = cb_Living4.SelectedItem.ToString();
                string _dependentPeopleKana5 = tb_DependentPeopleKana5.Text;
                string _dependentPeopleShimei5 = tb_DependentPeopleShimei5.Text;
                string _dependentPeopleBirth5 = CheckDateTime(dtp_DependentPeopleBirth5);
                string _relationship5 = tb_Relationship5.Text;
                string _living5 = cb_Living5.SelectedItem.ToString();
                string _dependentPeopleKana6 = tb_DependentPeopleKana6.Text;
                string _dependentPeopleShimei6 = tb_DependentPeopleShimei6.Text;
                string _dependentPeopleBirth6 = CheckDateTime(dtp_DependentPeopleBirth6);
                string _relationship6 = tb_Relationship6.Text;
                string _living6 = cb_Living6.SelectedItem.ToString();

                string _trainsportation1 = tb_Trainsportation1.Text;
                string _beginTrain1 = tb_BeginTrain1.Text;
                string _endTrain1 = tb_EndTrain1.Text;
                int _monthRegular1;
                bool result_monthRegular1 = int.TryParse(tb_MonthRegular1.Text,out _monthRegular1);
                string _trainsportation2 = tb_Trainsportation2.Text;
                string _beginTrain2 = tb_BeginTrain2.Text;
                string _endTrain2 = tb_EndTrain2.Text;
                int _monthRegular2;
                bool result_monthRegular2 = int.TryParse(tb_MonthRegular2.Text,out _monthRegular2);
                string _trainsportation3 = tb_Trainsportation3.Text;
                string _beginTrain3 = tb_BeginTrain3.Text;
                string _endTrain3 = tb_EndTrain3.Text;
                int _monthRegular3;
                bool result_monthRegular3 = int.TryParse(tb_MonthRegular3.Text,out _monthRegular3);
                string _trainsportation4 = tb_Trainsportation4.Text;
                string _beginTrain4 = tb_BeginTrain4.Text;
                string _endTrain4 = tb_EndTrain4.Text;
                int _monthRegular4;
                bool result_monthRegular4 = int.TryParse(tb_MonthRegular4.Text,out _monthRegular4);
                string _carkm = cb_Carkm.SelectedItem.ToString();
                int _carMoney;
                bool result_carMoney = int.TryParse(tb_CarMoney.Text,out _carMoney);
                int _totalMoneyTrans = _monthRegular1 + _monthRegular2 + _monthRegular3 + _monthRegular4;
               
                string _reason = string.Empty;
                string _changeDateFrom = string.Empty;
                string _changeDate = string.Empty;
                double _genkaritsu;
                bool result_genkaritsu = double.TryParse("", out _genkaritsu);
                int _teateGaku;
                bool result_teateGaku = int.TryParse("", out _teateGaku);
                string _accountCode = string.Empty;
                int _chingin = _basicSalary + _seikinTeate + _gaikinTeate + _gijutsuTeate + _shikakuTeate + _yakushokuTeate + _eigyoTeate;;
                
                string _chinginType = string.Empty;
                int _kyuyoKojoGaku ;
                bool result_kyuyoKojoGaku = int.TryParse("", out _kyuyoKojoGaku);
                int _workTime ;
                bool result_workTime = int.TryParse("", out _workTime);
                string _teateType = string.Empty;

                dto_allInfo = new DTO_AllInfor(_idCode, _romaji, _furigana, _sex, _birth, _nationality,
                        _inCompanyDate, _cardType, _cardTime, _cardTimeOut, _outTime, _companyCode, _companyName, _workType,
                        _closingDate, _zipCode, _address1, _address2, _address3, _address4, _address5, _mobliePhone, _phone, _createPeople, _position, _hakenRyokin, _hakenRyokinType,
                        _shiharaiType,_tax,_salaryType,_basicSalary,_seikinTeate,_gaikinTeate,_gijutsuTeate,_shikakuTeate,
                        _yakushokuTeate,_eigyoTeate,_kazokuTeate,_jutakuTeate,_bekkyoTeate,_tsukinTeate,_park,_dormitoryFee,
                        _waterFee,_employStatus,_employTime1,_employTime2,_bankName,_bankNameType,_branchName,_branchNameType,
                        _accountName,_bankCode,_branchCode,_accountCode1,_accountCode2,_accountCode3,_accountCode4,_accountCode5,
                        _accountCode6,_accountCode7,_accountCode8,_travelType,_houseName,_room,_inHouseDate,_kouyouhokenDate,_shakaihokenDate,
                        _dependentPeople,_residentPeople,_healthInsurancePeople,_contractType,_contractRequire,_myCompany,
                        _workContent,_workTime1,_workTime2,_workTime3,_workTime4,_relaxTime,_insureCard,_pastCompany1,_nienhieu1,
                        _beginYear1,_beginMonth1,_endYear1,_endMonth1,_pastCompany2,_nienhieu2,_beginYear2,_beginMonth2,
                        _endYear2,_endMonth2,_pensionBook,_dependentPeopleKana1,_dependentPeopleShimei1,_dependentPeopleBirth1,
                        _relationship1,_living1,_dependentPeopleKana2,_dependentPeopleShimei2,_dependentPeopleBirth2,
                        _relationship2,_living2,_dependentPeopleKana3,_dependentPeopleShimei3,_dependentPeopleBirth3,
                        _relationship3,_living3,_dependentPeopleKana4,_dependentPeopleShimei4,_dependentPeopleBirth4,
                        _relationship4,_living4,_dependentPeopleKana5,_dependentPeopleShimei5,_dependentPeopleBirth5,
                        _relationship5,_living5,_dependentPeopleKana6,_dependentPeopleShimei6,_dependentPeopleBirth6,
                        _relationship6,_living6,_trainsportation1,_beginTrain1,_endTrain1,_monthRegular1,
                        _trainsportation2,_beginTrain2,_endTrain2,_monthRegular2,
                        _trainsportation3,_beginTrain3,_endTrain3,_monthRegular3,
                        _trainsportation4,_beginTrain4,_endTrain4,_monthRegular4,_carkm,_carMoney,_totalMoneyTrans,
                        _reason, _changeDateFrom, _changeDate, _genkaritsu, _teateGaku, _accountCode, _chingin, _chinginType, _kyuyoKojoGaku, _workTime,_teateType);

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

        //Xu ly combobox cho phep insert string empty to database
        public string HandleCombobox(ComboBox cbox, string text) {
            if (cbox.SelectedIndex != -1)
            {
                text = cb_Sex.SelectedItem.ToString();
            }
            else
            {
                text = string.Empty;
            }
            return text;
        }

        //Validate ko nhap name, age, sex
        private void tb_RomajiName_Validating(object sender, CancelEventArgs e)
        {
            bll_handleFunc.ValidateControls(tb_RomajiName, errorProvider);
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
            dtp_CardTimeStart.Format = DateTimePickerFormat.Long;
        }

        private void dtp_CardTimeOut_ValueChanged(object sender, EventArgs e)
        {
            dtp_CardTimeOver.Format = DateTimePickerFormat.Long;
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

        private void tb_Park_TextChanged(object sender, EventArgs e)
        {
            int park_money;
            bool result_tb_park = int.TryParse(tb_Park.Text, out park_money);
            int dormitoryFee;
            bool result_dormitoryFee = int.TryParse(tb_DormitoryFee.Text, out dormitoryFee);
            int water_money;
            bool result_water_money = int.TryParse(tb_WaterFee.Text, out water_money);
            lb_tongtienkhautru.Text = (park_money + dormitoryFee + water_money).ToString();
        }

        private void tb_DormitoryFee_TextChanged(object sender, EventArgs e)
        {
            int park_money;
            bool result_tb_park = int.TryParse(tb_Park.Text, out park_money);
            int dormitoryFee;
            bool result_dormitoryFee = int.TryParse(tb_DormitoryFee.Text, out dormitoryFee);
            int water_money;
            bool result_water_money = int.TryParse(tb_WaterFee.Text, out water_money);
            lb_tongtienkhautru.Text = (park_money + dormitoryFee + water_money).ToString();
        }

        private void tb_WaterFee_TextChanged(object sender, EventArgs e)
        {
            int park_money;
            bool result_tb_park = int.TryParse(tb_Park.Text, out park_money);
            int dormitoryFee;
            bool result_dormitoryFee = int.TryParse(tb_DormitoryFee.Text, out dormitoryFee);
            int water_money;
            bool result_water_money = int.TryParse(tb_WaterFee.Text, out water_money);
            lb_tongtienkhautru.Text = (park_money + dormitoryFee + water_money).ToString();
        }

        private void total_chingin() {
            int _basicSalary;
            bool result_basicSalary = int.TryParse(tb_BasicSalary.Text, out _basicSalary);
            int _seikinTeate;
            bool result_seikinTeate = int.TryParse(tb_SeikinTeate.Text, out _seikinTeate);
            int _gaikinTeate;
            bool result_gaikinTeate = int.TryParse(tb_GaikinTeate.Text, out _gaikinTeate);
            int _gijutsuTeate;
            bool result_gijutsuTeate = int.TryParse(tb_GijutsuTeate.Text, out _gijutsuTeate);
            int _shikakuTeate;
            bool result_shikakuTeate = int.TryParse(tb_ShikakuTeate.Text, out _shikakuTeate);
            int _yakushokuTeate;
            bool result_yakushokuTeate = int.TryParse(tb_YakushokuTeate.Text, out _yakushokuTeate);
            int _eigyoTeate;
            bool result_eigyoTeate = int.TryParse(tb_EigyoTeate.Text, out _eigyoTeate);
            lb_chingin.Text = string.Format("{0:n0}", (_basicSalary + _seikinTeate + _gaikinTeate + _gijutsuTeate + _shikakuTeate + _yakushokuTeate + _eigyoTeate));

        }
        private void tb_BasicSalary_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void tb_SeikinTeate_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void tb_GaikinTeate_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void tb_GijutsuTeate_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void tb_ShikakuTeate_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void tb_YakushokuTeate_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void tb_EigyoTeate_TextChanged(object sender, EventArgs e)
        {
            total_chingin();
        }

        private void cb_EmployStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_EmployStatus.SelectedIndex == 0)
            {
                tbl_employTime.Enabled = false;
                dtp_EmployTime1.CustomFormat = " ";
                dtp_EmployTime1.Format = DateTimePickerFormat.Custom;
                dtp_EmployTime2.CustomFormat = " ";
                dtp_EmployTime1.Format = DateTimePickerFormat.Custom;
            }
            else
            {
                tbl_employTime.Enabled = true;
            }
        }

    }
}
