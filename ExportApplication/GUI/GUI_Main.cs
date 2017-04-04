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

        //Double click vao mỗi Nhân viên
        private void dtGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
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

        private void bt_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bt_print_Click(object sender, EventArgs e)
        {
            string name = dtGridView.SelectedCells[0].Value.ToString();
            GUI_Print gui_print = new GUI_Print(name);
            gui_print.Show();
        }

        public delegate void delPassData(string text);
        private void btEdit_Click(object sender, EventArgs e)
        {
            GUI_EditOption gui_editoption = new GUI_EditOption();
            gui_editoption.Show();
            List<String> list = new List<String>();
            if (dtGridView.SelectedCells.Count > 0)
            {
                int selectedrowindex = dtGridView.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = dtGridView.Rows[selectedrowindex];
                string a = Convert.ToString(selectedRow.Cells["氏名"].Value);
                GUI_Edit edit = new GUI_Edit();
                delPassData del = new delPassData(edit.funData);
                del(a);

            }
        }

        private void bt_addNew_Click(object sender, EventArgs e)
        {
            GUI_AddNew gui_addnew = new GUI_AddNew();
            gui_addnew.Show();
        }
    }
}
