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

namespace ExportApplication
{
    public partial class GUI_Edit : Form
    {
        BLL_Edit bll_edit = new BLL_Edit();
        // Enable and disable nhung truong duoc chon de edit
        public void TakeThis(IList<int> list)
        {
            // TB_IDCode.Text = String.Join(Environment.NewLine, list);
            for (int i = 0; i < list.Count; i++)
            {
                switch (list[i])
                {
                    case 0:
                        TB_IDCode.Enabled = true;
                        label1.Font = new Font(label1.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 1:
                        TB_RomajiName.Enabled = true;
                        label2.Font = new Font(label2.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 2:
                        TB_FuriganaName.Enabled = true;
                        label3.Font = new Font(label3.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 3:
                        TB_CompanyCode.Enabled = true;
                        label4.Font = new Font(label4.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 4:
                        TB_CompanyName.Enabled = true;
                        label5.Font = new Font(label5.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 5:
                        CB_Sex.Enabled = true;
                        label6.Font = new Font(label6.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 6:
                        CB_ShiharaiType.Enabled = true;
                        label7.Font = new Font(label7.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 7:
                        CB_ClosingDate.Enabled = true;
                        label8.Font = new Font(label8.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 8:
                        DTP_Birth.Enabled = true;
                        label9.Font = new Font(label9.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 9:
                        TB_Reason.Enabled = true;
                        label10.Font = new Font(label10.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 10:
                        DTP_ChangeDate.Enabled = true;
                        label11.Font = new Font(label11.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 11:
                        DTP_ChangeDateFrom.Enabled = true;
                        label12.Font = new Font(label12.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 12:
                        TB_ZipCode.Enabled = true;
                        TB_Address.Enabled = true;
                        label24.Font = new Font(label24.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 13:
                        CB_TravelType.Enabled = true;
                        label22.Font = new Font(label22.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 14:
                        TB_KyuyoKojoGaku.Enabled = true;
                        label36.Font = new Font(label36.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 15:
                        TB_HakenRyokin.Enabled = true;
                        CB_HakenRyokinType.Enabled = true;
                        label27.Font = new Font(label27.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 16:
                        TB_TeateGaku.Enabled = true;
                        label34.Font = new Font(label34.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 17:
                        TB_TsukinTeate.Enabled = true;
                        label32.Font = new Font(label32.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 18:
                        TB_Chingin.Enabled = true;
                        CB_ChinginType.Enabled = true;
                        label30.Font = new Font(label30.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 19:
                        CB_WorkType.Enabled = true;
                        label15.Font = new Font(label15.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 20:
                        CB_Tax.Enabled = true;
                        label26.Font = new Font(label26.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 21:
                        CB_CardType.Enabled = true;
                        label20.Font = new Font(label20.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 22:
                        DTP_CardTimeStart.Enabled = true;
                        DTP_CardTimeOver.Enabled = true;
                        label14.Font = new Font(label14.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 23:
                        DTP_EmployTime1.Enabled = true;
                        DTP_EmployTime2.Enabled = true;
                        label21.Font = new Font(label21.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 24:
                        TB_WorkTime.Enabled = true;
                        label38.Font = new Font(label38.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 25:
                        TB_BankCode.Enabled = true;
                        label39.Font = new Font(label39.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 26:
                        TB_BranchCode.Enabled = true;
                        label40.Font = new Font(label40.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 27:
                        TB_BankName.Enabled = true;
                        CB_BankNameType.Enabled = true;
                        label41.Font = new Font(label41.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 28:
                        TB_BranchName.Enabled = true;
                        CB_BranchNameType.Enabled = true;
                        label42.Font = new Font(label42.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 29:
                        TB_AccountName.Enabled = true;
                        label43.Font = new Font(label43.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 30:
                        TB_AccountCode.Enabled = true;
                        label44.Font = new Font(label44.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 31:
                        DTP_KoyoHokenDate.Enabled = true;
                        label45.Font = new Font(label45.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 32:
                        DTP_CompanyInsureDate.Enabled = true;
                        label46.Font = new Font(label46.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 33:
                        TB_DependentPeople.Enabled = true;
                        label47.Font = new Font(label47.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 34:
                        TB_ResidentPeople.Enabled = true;
                        label48.Font = new Font(label48.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                    case 35:
                        TB_HealthInsurancePeople.Enabled = true;
                        label49.Font = new Font(label49.Font.Name, 9, FontStyle.Bold | FontStyle.Underline);
                        break;
                }
            }
        }

        public GUI_Edit()
        {
            InitializeComponent();

        }
        public static string name;
        public void funData(string text)
        {
            name = text;
        }
        private void GUI_Edit_Load(object sender, EventArgs e)
        {

            DataTable dt = bll_edit.EditForm();
            TB_IDCode.Text = name;
            //TB_IDCode.Text = dt.Rows[0].Field<string>("IDCode");
            TB_RomajiName.Text = dt.Rows[0].Field<string>("RomajiName");
            TB_FuriganaName.Text = dt.Rows[0].Field<string>("FuriganaName");
            TB_CompanyCode.Text = dt.Rows[0].Field<string>("CompanyCode");
            TB_CompanyName.Text = dt.Rows[0].Field<string>("CompanyName");
            CB_Sex.Text = dt.Rows[0].Field<string>("Sex");
            CB_ShiharaiType.Text = dt.Rows[0].Field<string>("ShiharaiType");
            CB_Tax.Text = dt.Rows[0].Field<string>("Tax");
            DTP_Birth.Text = dt.Rows[0].Field<string>("Birth");
            TB_Reason.Text = dt.Rows[0].Field<string>("Reason");
            if (dt.Rows[0].Field<string>("ChangeDate") == null)
            {
                DTP_ChangeDate.Format = DateTimePickerFormat.Custom;
                DTP_ChangeDate.CustomFormat = "    ";
            }
            else
            {
                DTP_ChangeDate.Text = dt.Rows[0].Field<string>("ChangeDate");
            }
            if (dt.Rows[0].Field<string>("ChangeDateFrom") == null)
            {
                DTP_ChangeDateFrom.Format = DateTimePickerFormat.Custom;
                DTP_ChangeDateFrom.CustomFormat = "    ";
            }
            else
            {
                DTP_ChangeDateFrom.Text = dt.Rows[0].Field<string>("ChangeDateFrom");
            }

            TB_ZipCode.Text = dt.Rows[0].Field<string>("ZipCode");
            TB_ZipCode1.Text = dt.Rows[0].Field<string>("ZipCode");
            TB_Address.Text = dt.Rows[0].Field<string>("Address");
            TB_Address1.Text = dt.Rows[0].Field<string>("Address");
            CB_TravelType.Text = dt.Rows[0].Field<string>("TravelType");
            if (dt.Rows[0].Field<string>("EmployTime1") == null)
            {
                DTP_EmployTime1.Format = DateTimePickerFormat.Custom;
                DTP_EmployTime1.CustomFormat = "    ";
            }
            else
            {
                DTP_EmployTime1.Text = dt.Rows[0].Field<string>("EmployTime1");
            }

            if (dt.Rows[0].Field<string>("EmployTime2") == null)
            {
                DTP_EmployTime2.Format = DateTimePickerFormat.Custom;
                DTP_EmployTime2.CustomFormat = "    ";
            }
            else
            {
                DTP_EmployTime2.Text = dt.Rows[0].Field<string>("EmployTime1");
            }

            CB_CardType.Text = dt.Rows[0].Field<string>("CardType");
            if (dt.Rows[0].Field<string>("CardTimeOver") == null)
            {
                DTP_CardTimeOver.Format = DateTimePickerFormat.Custom;
                DTP_CardTimeOver.CustomFormat = "    ";
            }
            else
            {
                DTP_CardTimeOver.Text = dt.Rows[0].Field<string>("CardTimeOver");
            }

            if (dt.Rows[0].Field<string>("CardTimeStart") == null)
            {
                DTP_CardTimeStart.Format = DateTimePickerFormat.Custom;
                DTP_CardTimeStart.CustomFormat = "    ";
            }
            else
            {
                DTP_CardTimeStart.Text = dt.Rows[0].Field<string>("CardTimeStart");
            }

            CB_WorkType.Text = dt.Rows[0].Field<string>("WorkType");
            CB_ClosingDate.Text = dt.Rows[0].Field<string>("ClosingDate");
            TB_HakenRyokin.Text = dt.Rows[0].Field<string>("HakenRyokin");
            CB_HakenRyokinType.Text = dt.Rows[0].Field<string>("HakenRyokinType");
            TB_Chingin.Text = dt.Rows[0].Field<string>("Chingin");
            CB_ChinginType.Text = dt.Rows[0].Field<string>("ChinginType");
            TB_TsukinTeate.Text = dt.Rows[0].Field<string>("TsukinTeate");
            TB_TeateGaku.Text = dt.Rows[0].Field<string>("TeateGaku");
            TB_KyuyoKojoGaku.Text = dt.Rows[0].Field<string>("KyuyoKojoGaku");
            TB_WorkTime.Text = dt.Rows[0].Field<string>("WorkTime");
            TB_BankName.Text = dt.Rows[0].Field<string>("BankName");
            CB_BankNameType.Text = dt.Rows[0].Field<string>("BankNameType");
            TB_BranchName.Text = dt.Rows[0].Field<string>("BranchName");
            CB_BranchNameType.Text = dt.Rows[0].Field<string>("BranchNameType");
            TB_AccountName.Text = dt.Rows[0].Field<string>("AccountName");
            TB_BankCode.Text = dt.Rows[0].Field<string>("BankCode");
            TB_BranchCode.Text = dt.Rows[0].Field<string>("BranchCode");
            TB_AccountCode.Text = dt.Rows[0].Field<string>("AccountCode");

            if (dt.Rows[0].Field<string>("CompanyInsureDate") == null)
            {
                DTP_CompanyInsureDate.Format = DateTimePickerFormat.Custom;
                DTP_CompanyInsureDate.CustomFormat = "    ";
            }
            else
            {
                DTP_CompanyInsureDate.Text = dt.Rows[0].Field<string>("CompanyInsureDate");
            }

            if (dt.Rows[0].Field<string>("KoyoHokenDate") == null)
            {
                DTP_KoyoHokenDate.Format = DateTimePickerFormat.Custom;
                DTP_KoyoHokenDate.CustomFormat = "    ";
            }
            else
            {
                DTP_KoyoHokenDate.Text = dt.Rows[0].Field<string>("KoyoHokenDate");
            }

            TB_DependentPeople.Text = dt.Rows[0].Field<Int32>("DependentPeople").ToString();
            TB_ResidentPeople.Text = dt.Rows[0].Field<Int32>("ResidentPeople").ToString();
            TB_HealthInsurancePeople.Text = dt.Rows[0].Field<Int32>("HealthInsurancePeople").ToString();

        }
    }
}
