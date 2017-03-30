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
    public partial class GUI_Print : Form
    {
        public static string setName;
        public GUI_Print(string getName)
        {
            InitializeComponent();
            setName = getName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("印刷を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        MessageBox.Show("0. Hello "+setName);
                        break;
                    case 1:
                        MessageBox.Show("1");
                        break;
                    case 2:
                        MessageBox.Show("2");
                        break;
                    case 3:
                        MessageBox.Show("3");
                        break;

                
                }

            }
            else if (dialogResult == DialogResult.No)
            {
            }
        }
    }
}
