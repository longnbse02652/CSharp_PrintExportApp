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
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // clbEditOption
            // 
            this.clbEditOption.BorderStyle = System.Windows.Forms.BorderStyle.None;
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
            this.clbEditOption.Location = new System.Drawing.Point(0, 43);
            this.clbEditOption.MultiColumn = true;
            this.clbEditOption.Name = "clbEditOption";
            this.clbEditOption.Size = new System.Drawing.Size(392, 224);
            this.clbEditOption.TabIndex = 0;
            this.clbEditOption.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clbEditOption_MouseDown);
            // 
            // btNext
            // 
            this.btNext.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(168)))), ((int)(((byte)(171)))));
            this.btNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btNext.ForeColor = System.Drawing.Color.White;
            this.btNext.Location = new System.Drawing.Point(198, 273);
            this.btNext.Name = "btNext";
            this.btNext.Size = new System.Drawing.Size(194, 47);
            this.btNext.TabIndex = 1;
            this.btNext.Text = "次へ";
            this.btNext.UseVisualStyleBackColor = false;
            this.btNext.Click += new System.EventHandler(this.btNext_Click);
            // 
            // btCancel
            // 
            this.btCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(168)))), ((int)(((byte)(171)))));
            this.btCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btCancel.ForeColor = System.Drawing.Color.White;
            this.btCancel.Location = new System.Drawing.Point(1, 273);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(191, 47);
            this.btCancel.TabIndex = 2;
            this.btCancel.Text = "戻る";
            this.btCancel.UseVisualStyleBackColor = false;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(168)))), ((int)(((byte)(171)))));
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(392, 37);
            this.panel1.TabIndex = 3;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS PMincho", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(92, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(251, 15);
            this.label1.TabIndex = 4;
            this.label1.Text = "変更してほしい項目を選んでください！";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // GUI_EditOption
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(233)))), ((int)(((byte)(233)))));
            this.ClientSize = new System.Drawing.Size(392, 320);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btNext);
            this.Controls.Add(this.clbEditOption);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "GUI_EditOption";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "GUI_EditOption";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbEditOption;
        private System.Windows.Forms.Button btNext;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;

    }
}