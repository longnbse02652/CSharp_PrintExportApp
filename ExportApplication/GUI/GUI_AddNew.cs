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
        
        public GUI_AddNew()
        {
            InitializeComponent();
        }
        
        private void GUI_AddNew_Load(object sender, EventArgs e)
        {
            this.ActiveControl = tb_IDCode;
        }

        //nút này để lưu dữ liệu vào database
        private void bt_Save_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("登録を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string _idCode = tb_IDCode.Text;
                string _romaji = tb_RomajiName.Text;
                string _furigana = tb_FuriganaName.Text;
                string _sex = cb_Sex.SelectedItem.ToString();
                int _age = Int32.Parse(tb_Age.Text); 
                string _birth = bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtp_Birth);
                string _nationality = tb_Nationality.Text;
                string _inCompanyDate = bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtp_InCompanyDate);
                string _cardType = cb_CardType.SelectedItem.ToString();
                string _cardTime = bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtp_CardTime);
                string _cardTimeOut = bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtp_CardTimeOut);
                string _outTime = cb_OutTime.SelectedItem.ToString();
                string _companyCode = tb_CompanyCode.Text;
                string _companyName = tb_CompanyName.Text;
                string _workType = cb_WorkType.SelectedItem.ToString();
                string _closingDate = cb_ClosingDate.SelectedItem.ToString();
                int _zipCode = Int32.Parse(tb_ZipCode.Text);
                string _address = tb_Address.Text;
                string _mobliePhone = tb_MobliePhone.Text;
                string _phone = tb_Phone.Text;
                string _createPeople = tb_CreatePeople.Text;
                string _position = tb_Position.Text;

                DTO_AllInfor dto_allInfo = new DTO_AllInfor(_idCode, _romaji, _furigana, _sex, _age, _birth, _nationality,
                    _inCompanyDate, _cardType, _cardTime, _cardTimeOut, _outTime, _companyCode, _companyName, _workType,
                    _closingDate, _zipCode, _address, _mobliePhone, _phone, _createPeople, _position);

                if (bll_allInfo.Insert(dto_allInfo))
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

        private void bt_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
