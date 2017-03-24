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
            DataTable dt = bll_infor.GetToListView();
            dtGridView.DataSource = dt;
        }

       
        private void bt_addNew_Click(object sender, EventArgs e)
        {
            GUI_AddNew gui_addnew = new GUI_AddNew();
            gui_addnew.Show();
        }

        private void bt_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        public void ResetForm() {
            this.Invalidate();
        }
        
    }
}
