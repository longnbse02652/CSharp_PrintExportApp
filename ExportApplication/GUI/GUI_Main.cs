using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BLL;

namespace ExportApplication
{
    public partial class Main : Form
    {
        BLL_Infor bll_infor = new BLL_Infor();

        public Main()
        {
            InitializeComponent();
        }

        //hàm này để load dữ liệu lên ListView khi run system
        private void Main_Load(object sender, EventArgs e)
        {
            DataTable dt = bll_infor.GetToListView();
            foreach (DataRow row in dt.Rows)
            {
                ListViewItem item = new ListViewItem(row["RomajiName"].ToString());
                item.SubItems.Add(row["FuriganaName"].ToString());
                item.SubItems.Add(ConvertJapaneseCalendar(row["Birth"].ToString()));

                listView1.Items.Add(item);
            }

        }

        //hàm này là để chuyển định dang datime thông thường sang datetime của Nhật
        public string ConvertJapaneseCalendar(string datetime)
        {
            //string datetime = "1993/10/14";
            DateTime dt = Convert.ToDateTime(datetime);
            JapaneseCalendar myCal = new JapaneseCalendar();

            switch (myCal.GetEra(dt).ToString())
            {
                case "1":
                    datetime = "明治" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
                case "2":
                    datetime = "大正" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
                case "3":
                    datetime = "昭和" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
                case "4":
                    datetime = "平成" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
            }
            return datetime;
        }

        private void bt_addNew_Click(object sender, EventArgs e)
        {
            GUI_AddNew gui_addnew = new GUI_AddNew();
            gui_addnew.Show();
        }
    }
}
