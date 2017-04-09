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
using System.Runtime.InteropServices;

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
        public void LoadGridView()
        {
            DataTable dt = bll_infor.GetToListView();
            dtGridView.DataSource = dt;
        }

        //Double click vao mỗi Nhân viên
        private void dtGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string name = dtGridView.SelectedCells[0].Value.ToString();
            //MessageBox.Show(name);
            GUI_View gui_view = new GUI_View(name);
            gui_view.Show();

        }

        private void bt_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
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

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void tableLayoutPanel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }

        private void GUI_Main_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }

        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); 
        }
    }
}
