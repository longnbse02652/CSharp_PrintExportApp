namespace ExportApplication
{
    partial class GUI_Main
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
            this.bt_addNew = new System.Windows.Forms.Button();
            this.btEdit = new System.Windows.Forms.Button();
            this.bt_Exit = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dtGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dtGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // bt_addNew
            // 
            this.bt_addNew.Location = new System.Drawing.Point(43, 136);
            this.bt_addNew.Name = "bt_addNew";
            this.bt_addNew.Size = new System.Drawing.Size(86, 44);
            this.bt_addNew.TabIndex = 0;
            this.bt_addNew.Text = "新入登録";
            this.bt_addNew.UseVisualStyleBackColor = true;
            this.bt_addNew.Click += new System.EventHandler(this.bt_addNew_Click);
            // 
            // btEdit
            // 
            this.btEdit.Location = new System.Drawing.Point(43, 198);
            this.btEdit.Name = "btEdit";
            this.btEdit.Size = new System.Drawing.Size(86, 37);
            this.btEdit.TabIndex = 2;
            this.btEdit.Text = "編集";
            this.btEdit.UseVisualStyleBackColor = true;
            this.btEdit.Click += new System.EventHandler(this.btEdit_Click);
            // 
            // bt_Exit
            // 
            this.bt_Exit.Location = new System.Drawing.Point(43, 336);
            this.bt_Exit.Name = "bt_Exit";
            this.bt_Exit.Size = new System.Drawing.Size(86, 37);
            this.bt_Exit.TabIndex = 3;
            this.bt_Exit.Text = "終了";
            this.bt_Exit.UseVisualStyleBackColor = true;
            this.bt_Exit.Click += new System.EventHandler(this.bt_Exit_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(213, 85);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(311, 19);
            this.textBox1.TabIndex = 4;
            // 
            // dtGridView
            // 
            this.dtGridView.AllowUserToAddRows = false;
            this.dtGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtGridView.Location = new System.Drawing.Point(190, 136);
            this.dtGridView.Name = "dtGridView";
            this.dtGridView.ReadOnly = true;
            this.dtGridView.RowTemplate.Height = 21;
            this.dtGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dtGridView.Size = new System.Drawing.Size(334, 237);
            this.dtGridView.TabIndex = 5;
            this.dtGridView.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dtGridView_CellMouseDoubleClick);
            // 
            // GUI_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 424);
            this.Controls.Add(this.dtGridView);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.bt_Exit);
            this.Controls.Add(this.btEdit);
            this.Controls.Add(this.bt_addNew);
            this.Name = "GUI_Main";
            this.Text = "社員管理システム";
            this.Load += new System.EventHandler(this.Main_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bt_addNew;
        private System.Windows.Forms.Button btEdit;
        private System.Windows.Forms.Button bt_Exit;
        private System.Windows.Forms.TextBox textBox1;
        public System.Windows.Forms.DataGridView dtGridView;
    }
}

