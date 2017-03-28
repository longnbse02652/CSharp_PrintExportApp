namespace ExportApplication
{
    partial class GUI_EditOption
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.clbEditOption = new System.Windows.Forms.CheckedListBox();
            this.btNext = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clbEditOption
            // 
            this.clbEditOption.CheckOnClick = true;
            this.clbEditOption.FormattingEnabled = true;
            this.clbEditOption.Items.AddRange(new object[] {
            "社員ＣＤ",
            "氏名",
            "カナ",
            "企業ＣＤ",
            "企業名",
            "性別",
            "支払",
            "締日",
            "生年月日",
            "変更理由",
            "変更日",
            "変更締日",
            "住民票住所",
            "通勤形態",
            "給与控除額",
            "派遣・請金①",
            "手当額①",
            "通勤手当",
            "賃金",
            "就労形態",
            "税適",
            "在留資格",
            "在留期間",
            "雇用期間",
            "労働時間",
            "銀行CD",
            "支店CD",
            "銀行名",
            "支店名",
            "口座名義（カナ）",
            "口座番号",
            "雇用保険",
            "社会保険",
            "所得扶養数",
            "住民扶養数",
            "健保扶養数"});
            this.clbEditOption.Location = new System.Drawing.Point(0, 0);
            this.clbEditOption.MultiColumn = true;
            this.clbEditOption.Name = "clbEditOption";
            this.clbEditOption.Size = new System.Drawing.Size(468, 259);
            this.clbEditOption.TabIndex = 0;
            // 
            // btNext
            // 
            this.btNext.Location = new System.Drawing.Point(89, 267);
            this.btNext.Name = "btNext";
            this.btNext.Size = new System.Drawing.Size(87, 31);
            this.btNext.TabIndex = 1;
            this.btNext.Text = "次へ";
            this.btNext.UseVisualStyleBackColor = true;
            this.btNext.Click += new System.EventHandler(this.btNext_Click);
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(267, 267);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(87, 31);
            this.btCancel.TabIndex = 2;
            this.btCancel.Text = "キャンセル";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // GUI_EditOption
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 302);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btNext);
            this.Controls.Add(this.clbEditOption);
            this.Name = "GUI_EditOption";
            this.Text = "GUI_EditOption";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbEditOption;
        private System.Windows.Forms.Button btNext;
        private System.Windows.Forms.Button btCancel;

    }
}