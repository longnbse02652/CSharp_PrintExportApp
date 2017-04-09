using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
            //Lay tung gia tri da duoc selected trong checkbox in panel
            IList<string> list = new List<string>();
            foreach (Control c in panel2.Controls)
            {
                if ((c is CheckBox) && ((CheckBox)c).Checked)
                    list.Add(c.Text);
            }
            GUI_Edit gui_edit = new GUI_Edit();
            gui_edit.TakeThis(list);
            gui_edit.Show();
            this.Hide();
        }
        private void btCancel_Click(object sender, EventArgs e)
        {
            this.Close();
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
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();


        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }
    }
}
