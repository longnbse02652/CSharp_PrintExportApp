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
    public partial class GUI_EditOption : Form
    {
        public GUI_EditOption()
        {
            InitializeComponent();
        }

        public void btNext_Click(object sender, EventArgs e)
        {
            // this.Hide();
            // Lay tung gia tri da duoc selected trong listcheckbox
            IList<int> list = new List<int>();
            for (int i = 0; i < clbEditOption.CheckedIndices.Count; i++)
            {
                list.Add(clbEditOption.CheckedIndices[i]);
            }
            // Khai bao va get list sang form edit
            GUI_Edit gui_edit = new GUI_Edit();
            gui_edit.TakeThis(list);
            gui_edit.Show();
        }
        public void TakeThis(IList<int> list)
        {

        }
        private void btCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
