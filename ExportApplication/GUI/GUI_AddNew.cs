using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportApplication
{
    public partial class GUI_AddNew : Form
    {
        public GUI_AddNew()
        {
            InitializeComponent();
            
        }

        private void GUI_AddNew_Load(object sender, EventArgs e)
        {
            this.ActiveControl = tb_IDCode;
            
        }

        private void bt_Save_Click(object sender, EventArgs e)
        {

        }



    }
}
