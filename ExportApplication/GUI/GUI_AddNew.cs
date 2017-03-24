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

namespace ExportApplication
{
    public partial class GUI_AddNew : Form
    {
        BLL_AllInfor bll_allInfo = new BLL_AllInfor();
        BLL_HandleFunc bll_handleFunc = new BLL_HandleFunc(); //tạo object của class này ra để format trước khi chuyển xuống database
        
        public GUI_AddNew()
        {
            InitializeComponent();

        }
        
        private void GUI_AddNew_Load(object sender, EventArgs e)
        {
            this.ActiveControl = tb_IDCode;
        }

        //nút này để lưu dữ liệu vào database
        private void bt_Save_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("登録を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string romaji = tb_RomajiName.Text;
                string furigana = tb_FuriganaName.Text;
                string birth = bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtpicker_Birth);

                DTO_AllInfor dto_allInfo = new DTO_AllInfor(romaji, furigana, birth);
                if (bll_allInfo.Insert(dto_allInfo))
                {
                    MessageBox.Show("登録しました。");
                    this.Close();
                    GUI_Main main = new GUI_Main();
                    main.dtGridView.Refresh();
                }
                else
                {
                    MessageBox.Show("登録は失敗しました。");
                }
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //MessageBox.Show(bll_handleFunc.ConvertFromDatetimePicker_ToYYMMDD(dtpicker_Birth) + tb_FuriganaName.Text + tb_RomajiName.Text);
                //cancel Save 
            }

        }

        private void bt_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }



    }
}
