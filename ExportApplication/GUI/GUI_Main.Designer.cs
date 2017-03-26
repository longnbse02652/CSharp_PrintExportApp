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
            this.button2 = new System.Windows.Forms.Button();
            this.bt_Exit = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dtGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dtGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // bt_addNew
            // 
            this.bt_addNew.Location = new System.Drawing.Point(43, 147);
            this.bt_addNew.Name = "bt_addNew";
            this.bt_addNew.Size = new System.Drawing.Size(86, 48);
            this.bt_addNew.TabIndex = 0;
            this.bt_addNew.Text = "新入登録";
            this.bt_addNew.UseVisualStyleBackColor = true;
            this.bt_addNew.Click += new System.EventHandler(this.bt_addNew_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(43, 215);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(86, 40);
            this.button2.TabIndex = 2;
            this.button2.Text = "編集";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // bt_Exit
            // 
            this.bt_Exit.Location = new System.Drawing.Point(43, 364);
            this.bt_Exit.Name = "bt_Exit";
            this.bt_Exit.Size = new System.Drawing.Size(86, 40);
            this.bt_Exit.TabIndex = 3;
            this.bt_Exit.Text = "終了";
            this.bt_Exit.UseVisualStyleBackColor = true;
            this.bt_Exit.Click += new System.EventHandler(this.bt_Exit_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(213, 92);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(311, 20);
            this.textBox1.TabIndex = 4;
            // 
            // dtGridView
            // 
            this.dtGridView.AllowUserToAddRows = false;
            this.dtGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtGridView.Location = new System.Drawing.Point(190, 147);
            this.dtGridView.Name = "dtGridView";
            this.dtGridView.ReadOnly = true;
            this.dtGridView.RowTemplate.Height = 21;
            this.dtGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dtGridView.Size = new System.Drawing.Size(334, 257);
            this.dtGridView.TabIndex = 5;
            this.dtGridView.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dtGridView_CellMouseDoubleClick);
            // 
            // GUI_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 459);
            this.Controls.Add(this.dtGridView);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.bt_Exit);
            this.Controls.Add(this.button2);
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
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button bt_Exit;
        private System.Windows.Forms.TextBox textBox1;
        public System.Windows.Forms.DataGridView dtGridView;
    }
}

