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

namespace ExportApplication
{
    public partial class GUI_Dependent : Form
    {
        public GUI_Dependent()
        {
            InitializeComponent();
        }
        BLL_HandleFunc bll_handleFunc = new BLL_HandleFunc();
        BLL_Dependent bll_dependent = new BLL_Dependent();
        DataTable dt = new DataTable();
        public GUI_Dependent(string name)
        {
            InitializeComponent();
            nameRomaji = name;
            // MessageBox.Show(nameRomaji);
        }

        public static string nameRomaji;

        private void GUI_Dependent_Load(object sender, EventArgs e)
        {
            dt = bll_dependent.DependentForm(nameRomaji);
            foreach (DataRow row in dt.Rows)
            {
                tb_DependentPeopleKana1.Text = dt.Rows[0].Field<string>("DependentPeopleKana1");
                tb_DependentPeopleShimei1.Text = dt.Rows[0].Field<string>("DependentPeopleShimei1");
                tb_DependentPeopleKana2.Text = dt.Rows[0].Field<string>("DependentPeopleKana2");
                tb_DependentPeopleShimei2.Text = dt.Rows[0].Field<string>("DependentPeopleShimei2");
                tb_DependentPeopleKana3.Text = dt.Rows[0].Field<string>("DependentPeopleKana3");
                tb_DependentPeopleShimei3.Text = dt.Rows[0].Field<string>("DependentPeopleShimei3");
                tb_DependentPeopleKana4.Text = dt.Rows[0].Field<string>("DependentPeopleKana4");
                tb_DependentPeopleShimei4.Text = dt.Rows[0].Field<string>("DependentPeopleShimei4");
                tb_DependentPeopleKana5.Text = dt.Rows[0].Field<string>("DependentPeopleKana5");
                tb_DependentPeopleShimei5.Text = dt.Rows[0].Field<string>("DependentPeopleShimei5");
                tb_DependentPeopleKana6.Text = dt.Rows[0].Field<string>("DependentPeopleKana6");
                tb_DependentPeopleShimei6.Text = dt.Rows[0].Field<string>("DependentPeopleShimei6");
                tb_Relationship1.Text = dt.Rows[0].Field<string>("Relationship1");
                tb_Relationship2.Text = dt.Rows[0].Field<string>("Relationship2");
                tb_Relationship3.Text = dt.Rows[0].Field<string>("Relationship3");
                tb_Relationship4.Text = dt.Rows[0].Field<string>("Relationship4");
                tb_Relationship5.Text = dt.Rows[0].Field<string>("Relationship5");
                tb_Relationship6.Text = dt.Rows[0].Field<string>("Relationship6");
                if (string.IsNullOrEmpty(row["DependentPeopleBirth1"].ToString()) || row["DependentPeopleBirth1"].ToString() == " ")
                {
                    dtp_DependentPeopleBirth1.Format = DateTimePickerFormat.Custom;
                    dtp_DependentPeopleBirth1.CustomFormat = " ";
                }
                else { dtp_DependentPeopleBirth1.Text = dt.Rows[0].Field<string>("DependentPeopleBirth1"); }

                if (string.IsNullOrEmpty(row["DependentPeopleBirth2"].ToString()) || row["DependentPeopleBirth2"].ToString() == " ")
                {
                    dtp_DependentPeopleBirth2.Format = DateTimePickerFormat.Custom;
                    dtp_DependentPeopleBirth2.CustomFormat = " ";
                }
                else { dtp_DependentPeopleBirth2.Text = dt.Rows[0].Field<string>("DependentPeopleBirth2"); }

                if (string.IsNullOrEmpty(row["DependentPeopleBirth3"].ToString()) || row["DependentPeopleBirth3"].ToString() == " ")
                {
                    dtp_DependentPeopleBirth3.Format = DateTimePickerFormat.Custom;
                    dtp_DependentPeopleBirth3.CustomFormat = " ";
                }
                else { dtp_DependentPeopleBirth3.Text = dt.Rows[0].Field<string>("DependentPeopleBirth3"); }

                if (string.IsNullOrEmpty(row["DependentPeopleBirth4"].ToString()) || row["DependentPeopleBirth4"].ToString() == " ")
                {
                    dtp_DependentPeopleBirth4.Format = DateTimePickerFormat.Custom;
                    dtp_DependentPeopleBirth4.CustomFormat = " ";
                }
                else { dtp_DependentPeopleBirth4.Text = dt.Rows[0].Field<string>("DependentPeopleBirth4"); }

                if (string.IsNullOrEmpty(row["DependentPeopleBirth5"].ToString()) || row["DependentPeopleBirth5"].ToString() == " ")
                {
                    dtp_DependentPeopleBirth5.Format = DateTimePickerFormat.Custom;
                    dtp_DependentPeopleBirth5.CustomFormat = " ";
                }
                else { dtp_DependentPeopleBirth5.Text = dt.Rows[0].Field<string>("DependentPeopleBirth5"); }

                if (string.IsNullOrEmpty(row["DependentPeopleBirth6"].ToString()) || row["DependentPeopleBirth6"].ToString() == " ")
                {
                    dtp_DependentPeopleBirth6.Format = DateTimePickerFormat.Custom;
                    dtp_DependentPeopleBirth6.CustomFormat = " ";
                }
                else { dtp_DependentPeopleBirth6.Text = dt.Rows[0].Field<string>("DependentPeopleBirth6"); }

                cb_Living1.Text = dt.Rows[0].Field<string>("Living1");
                cb_Living2.Text = dt.Rows[0].Field<string>("Living2");
                cb_Living3.Text = dt.Rows[0].Field<string>("Living3");
                cb_Living4.Text = dt.Rows[0].Field<string>("Living4");
                cb_Living5.Text = dt.Rows[0].Field<string>("Living5");
                cb_Living6.Text = dt.Rows[0].Field<string>("Living6");
                object dependentpeople = row["DependentPeople"];
                object residentpeople = row["ResidentPeople"];
                object healthinsurancepeople = row["HealthInsurancePeople"];
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
            }
        }

        public DTO_Dependent Dependent()
        {

            string _DependentPeopleKana1 = tb_DependentPeopleKana1.Text;
            string _DependentPeopleKana2 = tb_DependentPeopleKana2.Text;
            string _DependentPeopleKana3 = tb_DependentPeopleKana3.Text;
            string _DependentPeopleKana4 = tb_DependentPeopleKana4.Text;
            string _DependentPeopleKana5 = tb_DependentPeopleKana5.Text;
            string _DependentPeopleKana6 = tb_DependentPeopleKana6.Text;

            string _DependentPeopleShimei1 = tb_DependentPeopleShimei1.Text;
            string _DependentPeopleShimei2 = tb_DependentPeopleShimei2.Text;
            string _DependentPeopleShimei3 = tb_DependentPeopleShimei3.Text;
            string _DependentPeopleShimei4 = tb_DependentPeopleShimei4.Text;
            string _DependentPeopleShimei5 = tb_DependentPeopleShimei5.Text;
            string _DependentPeopleShimei6 = tb_DependentPeopleShimei6.Text;

            string _Relationship1 = tb_Relationship1.Text;
            string _Relationship2 = tb_Relationship2.Text;
            string _Relationship3 = tb_Relationship3.Text;
            string _Relationship4 = tb_Relationship4.Text;
            string _Relationship5 = tb_Relationship5.Text;
            string _Relationship6 = tb_Relationship6.Text;
            string _DependentPeopleBirth1 = CheckDateTime(dtp_DependentPeopleBirth1);
            string _DependentPeopleBirth2 = CheckDateTime(dtp_DependentPeopleBirth2);
            string _DependentPeopleBirth3 = CheckDateTime(dtp_DependentPeopleBirth3);
            string _DependentPeopleBirth4 = CheckDateTime(dtp_DependentPeopleBirth4);
            string _DependentPeopleBirth5 = CheckDateTime(dtp_DependentPeopleBirth5);
            string _DependentPeopleBirth6 = CheckDateTime(dtp_DependentPeopleBirth6);
            string _RomajiName = nameRomaji;
            int _DependentPeople;
            bool result_DependentPeople = int.TryParse(TB_DependentPeople.Text, out _DependentPeople);
            int _ResidentPeople;
            bool result_ResidentPeople = int.TryParse(TB_ResidentPeople.Text, out _ResidentPeople);
            int _HealthInsurancePeople;
            bool result_HealthInsurancePeople = int.TryParse(TB_HealthInsurancePeople.Text, out _HealthInsurancePeople);
            string _Living1, _Living2, _Living3, _Living4, _Living5, _Living6;
            if (cb_Living1.SelectedIndex != -1) { _Living1 = cb_Living1.SelectedItem.ToString(); } else { _Living1 = string.Empty; }
            if (cb_Living2.SelectedIndex != -1) { _Living2 = cb_Living2.SelectedItem.ToString(); } else { _Living2 = string.Empty; }
            if (cb_Living3.SelectedIndex != -1) { _Living3 = cb_Living3.SelectedItem.ToString(); } else { _Living3 = string.Empty; }
            if (cb_Living4.SelectedIndex != -1) { _Living4 = cb_Living4.SelectedItem.ToString(); } else { _Living4 = string.Empty; }
            if (cb_Living5.SelectedIndex != -1) { _Living5 = cb_Living5.SelectedItem.ToString(); } else { _Living5 = string.Empty; }
            if (cb_Living6.SelectedIndex != -1) { _Living6 = cb_Living6.SelectedItem.ToString(); } else { _Living6 = string.Empty; }
            DTO_Dependent dto_dependent = new DTO_Dependent(_DependentPeopleKana1, _DependentPeopleKana2, _DependentPeopleKana3,
       _DependentPeopleKana4, _DependentPeopleKana5, _DependentPeopleKana6, _DependentPeopleShimei1, _DependentPeopleShimei2, _DependentPeopleShimei3, _DependentPeopleShimei4,
       _DependentPeopleShimei5, _DependentPeopleShimei6, _Relationship1, _Relationship2, _Relationship3, _Relationship4, _Relationship5, _Relationship6, _DependentPeopleBirth1,
       _DependentPeopleBirth2, _DependentPeopleBirth3, _DependentPeopleBirth4, _DependentPeopleBirth5, _DependentPeopleBirth6, _Living1, _Living2, _Living3, _Living4, _Living5, _Living6, _RomajiName,
       _DependentPeople, _ResidentPeople, _HealthInsurancePeople);
            return dto_dependent;
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
        private void BT_Save_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BT_Save_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("保存を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_dependent.Insert(Dependent()))
                {
                    MessageBox.Show("保存しました。");
                    GUI_Edit obj = (GUI_Edit)Application.OpenForms["GUI_Edit"];
                    obj.TB_DependentPeople.Text = TB_DependentPeople.Text;
                    obj.TB_ResidentPeople.Text = TB_ResidentPeople.Text;
                    obj.TB_HealthInsurancePeople.Text = TB_HealthInsurancePeople.Text;
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dtp_DependentPeopleBirth2_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth2.Format = DateTimePickerFormat.Long;
        }

        private void dtp_DependentPeopleBirth1_ValueChanged(object sender, EventArgs e)
        {
            dtp_DependentPeopleBirth1.Format = DateTimePickerFormat.Long;
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
