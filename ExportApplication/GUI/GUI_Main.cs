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
    public partial class GUI_Main : Form
    {
        BLL_AllInfor bll_infor = new BLL_AllInfor();
        BLL_HandleFunc bll_handleFunc = new BLL_HandleFunc();
       
        public GUI_Main()
        {
            InitializeComponent();
        }

        //hàm này để load dữ liệu lên ListView khi run system
        private void Main_Load(object sender, EventArgs e)
        {
            LoadGridView();
        }
        public void LoadGridView() {
            DataTable dt = bll_infor.GetToListView();
            dtGridView.DataSource = dt;
        }

        //Click Nút 新規登録
        private void bt_addNew_Click(object sender, EventArgs e)
        {
            GUI_AddNew gui_addnew = new GUI_AddNew();
            gui_addnew.Show();
        }

        //Click nút 終了
        private void bt_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Double click vao mỗi Nhân viên
        private void dtGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dtGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                MessageBox.Show(dtGridView.Rows[e.RowIndex].Cells["氏名"].Value.ToString());
                GUI_AddNew gui_view = new GUI_AddNew();

            }
        }


        //public delegate void delPassData(string text);

        private void btEdit_Click(object sender, EventArgs e)
        {
            GUI_EditOption gui_editoption = new GUI_EditOption();
            gui_editoption.Show();
   
        }

    }
}
