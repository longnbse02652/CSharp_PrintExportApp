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
    public partial class GUI_Travel : Form
    {
        BLL_Travel bll_travel = new BLL_Travel();
        DataTable dt = new DataTable();
        public GUI_Travel(string name)
        {
            InitializeComponent();
            nameRomaji = name;
            // MessageBox.Show(nameRomaji);
        }

        public static string nameRomaji;


        private void GUI_Travel_Load(object sender, EventArgs e)
        {
            dt = bll_travel.TravelForm(nameRomaji);
            foreach (DataRow row in dt.Rows)
            {
                tb_Trainsportation1.Text = dt.Rows[0].Field<string>("Trainsportation1");
                tb_BeginTrain1.Text = dt.Rows[0].Field<string>("BeginTrain1");
                tb_EndTrain1.Text = dt.Rows[0].Field<string>("EndTrain1");

                tb_Trainsportation2.Text = dt.Rows[0].Field<string>("Trainsportation2");
                tb_BeginTrain2.Text = dt.Rows[0].Field<string>("BeginTrain2");
                tb_EndTrain2.Text = dt.Rows[0].Field<string>("EndTrain2");

                tb_Trainsportation3.Text = dt.Rows[0].Field<string>("Trainsportation3");
                tb_BeginTrain3.Text = dt.Rows[0].Field<string>("BeginTrain3");
                tb_EndTrain3.Text = dt.Rows[0].Field<string>("EndTrain3");

                tb_Trainsportation4.Text = dt.Rows[0].Field<string>("Trainsportation4");
                tb_BeginTrain4.Text = dt.Rows[0].Field<string>("BeginTrain4");
                tb_EndTrain4.Text = dt.Rows[0].Field<string>("EndTrain4");


                object MonthRegular4 = row["MonthRegular4"];
                object MonthRegular3 = row["MonthRegular3"];
                object MonthRegular2 = row["MonthRegular2"];
                object MonthRegular1 = row["MonthRegular1"];
                object CarMoney = row["CarMoney"];
                object TotalMoneyTrans = row["TotalMoneyTrans"];

                if (MonthRegular4 == DBNull.Value)
                {
                    tb_MonthRegular4.Text = "";
                }
                else { tb_MonthRegular4.Text = dt.Rows[0].Field<int>("MonthRegular4").ToString(); }
                if (MonthRegular3 == DBNull.Value)
                {
                    tb_MonthRegular3.Text = "";
                }
                else { tb_MonthRegular3.Text = dt.Rows[0].Field<int>("MonthRegular3").ToString(); }
                if (MonthRegular2 == DBNull.Value)
                {
                    tb_MonthRegular2.Text = "";
                }
                else { tb_MonthRegular2.Text = dt.Rows[0].Field<int>("MonthRegular2").ToString(); }
                if (MonthRegular1 == DBNull.Value)
                {
                    tb_MonthRegular1.Text = "";
                }
                else { tb_MonthRegular1.Text = dt.Rows[0].Field<int>("MonthRegular1").ToString(); }

                if (CarMoney == DBNull.Value)
                {
                    tb_CarMoney.Text = "";
                }
                else { tb_CarMoney.Text = dt.Rows[0].Field<int>("CarMoney").ToString(); }

                if (TotalMoneyTrans == DBNull.Value)
                {
                    tb_TotalMoneyTrans.Text = "";
                }
                else { tb_TotalMoneyTrans.Text = dt.Rows[0].Field<int>("TotalMoneyTrans").ToString(); }
                cb_Carkm.Text = dt.Rows[0].Field<string>("Carkm");


            }
        }

        public DTO_Travel updateTravel()
        {

            string _Trainsportation1 = tb_Trainsportation1.Text;
            string _RomajiName = nameRomaji;
            string _BeginTrain1 = tb_BeginTrain1.Text;
            string _EndTrain1 = tb_EndTrain1.Text;
            string _Trainsportation2 = tb_Trainsportation2.Text;
            string _BeginTrain2 = tb_BeginTrain2.Text;
            string _EndTrain2 = tb_EndTrain2.Text;
            string _Trainsportation3 = tb_Trainsportation3.Text;
            string _BeginTrain3 = tb_BeginTrain3.Text;
            string _EndTrain3 = tb_EndTrain3.Text;
            string _Trainsportation4 = tb_Trainsportation4.Text;
            string _BeginTrain4 = tb_BeginTrain4.Text;
            string _EndTrain4 = tb_EndTrain4.Text;
            string _Carkm;
            if (cb_Carkm.SelectedIndex == -1) { _Carkm = ""; } else { _Carkm = cb_Carkm.SelectedItem.ToString(); }
            int _MonthRegular1;
            bool result_MonthRegular1 = int.TryParse(tb_MonthRegular1.Text, out _MonthRegular1);
            int _MonthRegular2;
            bool result_MonthRegular2 = int.TryParse(tb_MonthRegular2.Text, out _MonthRegular2);
            int _MonthRegular3;
            bool result_MonthRegular3 = int.TryParse(tb_MonthRegular3.Text, out _MonthRegular3);
            int _MonthRegular4;
            bool result_MonthRegular4 = int.TryParse(tb_MonthRegular4.Text, out _MonthRegular4);
            int _CarMoney;
            bool result_CarMoney = int.TryParse(tb_CarMoney.Text, out _CarMoney);
            int _TotalMoneyTrans;
            bool result_TotalMoneyTrans = int.TryParse(tb_TotalMoneyTrans.Text, out _TotalMoneyTrans);

            DTO_Travel dto_travel = new DTO_Travel(_Trainsportation1, _BeginTrain1, _EndTrain1,
       _Trainsportation2, _BeginTrain2, _EndTrain2, _Trainsportation3, _BeginTrain3, _EndTrain3,
       _Trainsportation4, _BeginTrain4, _EndTrain4, _MonthRegular1, _MonthRegular2, _MonthRegular3, _MonthRegular4,
       _Carkm, _CarMoney, _TotalMoneyTrans, _RomajiName);
            return dto_travel;
        }
        private void BT_Save_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("保存を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (bll_travel.Insert(updateTravel()))
                {
                    MessageBox.Show("保存しました。");
                    GUI_Edit obj = (GUI_Edit)Application.OpenForms["GUI_Edit"];
                    obj.TB_TsukinTeate.Text = total.ToString();
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

        // tu dong tinh toan khi nhap gia tien
        public int total = 0, total1, total2, total3, total4, total5;
        private void tb_MonthRegular1_TextChanged(object sender, EventArgs e)
        {
            if (tb_MonthRegular1.Text == "")
            {
                tb_MonthRegular1.Text = "";
            }
            else
            {
                int MonthRegular1 = int.Parse(tb_MonthRegular1.Text);
                total1 = MonthRegular1;
                total = total1 + total2 + total3 + total4 + total5;
                tb_TotalMoneyTrans.Text = total.ToString();
            }
        }
        private void tb_MonthRegular2_TextChanged_1(object sender, EventArgs e)
        {
            if (tb_MonthRegular2.Text == "")
            {
                tb_MonthRegular2.Text = "";
            }
            else
            {
                int MonthRegular2 = int.Parse(tb_MonthRegular2.Text);
                total2 = MonthRegular2;
                total = total1 + total2 + total3 + total4 + total5;
                tb_TotalMoneyTrans.Text = total.ToString();
            }
        }

        private void tb_MonthRegular3_TextChanged_1(object sender, EventArgs e)
        {
            if (tb_MonthRegular3.Text == "")
            {
                tb_MonthRegular3.Text = "";
            }
            else
            {
                int MonthRegular3 = int.Parse(tb_MonthRegular3.Text);
                total3 = MonthRegular3;
                total = total1 + total2 + total3 + total4 + total5;
                tb_TotalMoneyTrans.Text = total.ToString();
            }
        }

        private void tb_MonthRegular4_TextChanged_1(object sender, EventArgs e)
        {
            if (tb_MonthRegular4.Text == "")
            {
                tb_MonthRegular4.Text = "";
            }
            else
            {
                int MonthRegular4 = int.Parse(tb_MonthRegular4.Text);
                total4 = MonthRegular4;
                total = total1 + total2 + total3 + total4 + total5;
                tb_TotalMoneyTrans.Text = total.ToString();
            }
        }

        private void tb_CarMoney_TextChanged_1(object sender, EventArgs e)
        {
            if (tb_CarMoney.Text == "")
            {
                tb_CarMoney.Text = "";
            }
            else
            {
                int CarMoney = int.Parse(tb_CarMoney.Text);
                total5 = CarMoney;
                total = total1 + total2 + total3 + total4 + total5;
                tb_TotalMoneyTrans.Text = total.ToString();
            }
        }

        // cancel button click
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tb_MonthRegular1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void tb_MonthRegular2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void tb_MonthRegular3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void tb_MonthRegular4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void tb_CarMoney_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void tb_TotalMoneyTrans_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void tableLayoutPanel75_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }

        private void GUI_Travel_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }

        private void label157_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }



    }
}
