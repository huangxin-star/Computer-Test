namespace Computer_Test
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label1 = new System.Windows.Forms.Label();
            this.selectMappingTable = new System.Windows.Forms.Button();
            this.selectModelTable = new System.Windows.Forms.Button();
            this.transform = new System.Windows.Forms.Button();
            this.mapTextBox = new System.Windows.Forms.TextBox();
            this.machineTextBox = new System.Windows.Forms.TextBox();
            this.savaPathTextBox = new System.Windows.Forms.TextBox();
            this.savePathButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Time = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(516, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(352, 46);
            this.label1.TabIndex = 0;
            this.label1.Text = "三厂Spec測試軟件";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // selectMappingTable
            // 
            this.selectMappingTable.BackColor = System.Drawing.Color.Salmon;
            this.selectMappingTable.ForeColor = System.Drawing.SystemColors.Desktop;
            this.selectMappingTable.Location = new System.Drawing.Point(579, 164);
            this.selectMappingTable.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.selectMappingTable.Name = "selectMappingTable";
            this.selectMappingTable.Size = new System.Drawing.Size(165, 48);
            this.selectMappingTable.TabIndex = 1;
            this.selectMappingTable.Text = "選擇映射表總表";
            this.selectMappingTable.UseVisualStyleBackColor = false;
            this.selectMappingTable.Click += new System.EventHandler(this.selectMappingTable_Click);
            // 
            // selectModelTable
            // 
            this.selectModelTable.BackColor = System.Drawing.Color.Salmon;
            this.selectModelTable.Location = new System.Drawing.Point(579, 241);
            this.selectModelTable.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.selectModelTable.Name = "selectModelTable";
            this.selectModelTable.Size = new System.Drawing.Size(165, 53);
            this.selectModelTable.TabIndex = 2;
            this.selectModelTable.Text = "選擇Spec文件";
            this.selectModelTable.UseVisualStyleBackColor = false;
            this.selectModelTable.Click += new System.EventHandler(this.selectModelTable_Click);
            // 
            // transform
            // 
            this.transform.BackColor = System.Drawing.Color.Orange;
            this.transform.ForeColor = System.Drawing.SystemColors.Desktop;
            this.transform.Location = new System.Drawing.Point(294, 432);
            this.transform.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.transform.Name = "transform";
            this.transform.Size = new System.Drawing.Size(165, 53);
            this.transform.TabIndex = 3;
            this.transform.Text = "轉檔";
            this.transform.UseVisualStyleBackColor = false;
            this.transform.Click += new System.EventHandler(this.transform_click);
            // 
            // mapTextBox
            // 
            this.mapTextBox.Location = new System.Drawing.Point(144, 176);
            this.mapTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.mapTextBox.MaximumSize = new System.Drawing.Size(500, 500);
            this.mapTextBox.Name = "mapTextBox";
            this.mapTextBox.Size = new System.Drawing.Size(367, 28);
            this.mapTextBox.TabIndex = 4;
            this.mapTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // machineTextBox
            // 
            this.machineTextBox.Location = new System.Drawing.Point(144, 256);
            this.machineTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.machineTextBox.Name = "machineTextBox";
            this.machineTextBox.Size = new System.Drawing.Size(367, 28);
            this.machineTextBox.TabIndex = 6;
            this.machineTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // savaPathTextBox
            // 
            this.savaPathTextBox.Location = new System.Drawing.Point(144, 344);
            this.savaPathTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.savaPathTextBox.Name = "savaPathTextBox";
            this.savaPathTextBox.Size = new System.Drawing.Size(367, 28);
            this.savaPathTextBox.TabIndex = 7;
            this.savaPathTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // savePathButton
            // 
            this.savePathButton.BackColor = System.Drawing.Color.Salmon;
            this.savePathButton.Location = new System.Drawing.Point(579, 330);
            this.savePathButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.savePathButton.Name = "savePathButton";
            this.savePathButton.Size = new System.Drawing.Size(165, 53);
            this.savePathButton.TabIndex = 8;
            this.savePathButton.Text = "選擇存放路徑";
            this.savePathButton.UseVisualStyleBackColor = false;
            this.savePathButton.Click += new System.EventHandler(this.savePath_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(138, 100);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 37);
            this.label2.TabIndex = 9;
            this.label2.Text = "路徑";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(59, 35);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(216, 60);
            this.pictureBox1.TabIndex = 31;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(576, 496);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 18);
            this.label3.TabIndex = 32;
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // Time
            // 
            this.Time.AutoSize = true;
            this.Time.Location = new System.Drawing.Point(616, 496);
            this.Time.Name = "Time";
            this.Time.Size = new System.Drawing.Size(0, 18);
            this.Time.TabIndex = 33;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(900, 540);
            this.Controls.Add(this.Time);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.savePathButton);
            this.Controls.Add(this.savaPathTextBox);
            this.Controls.Add(this.machineTextBox);
            this.Controls.Add(this.mapTextBox);
            this.Controls.Add(this.transform);
            this.Controls.Add(this.selectModelTable);
            this.Controls.Add(this.selectMappingTable);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button selectMappingTable;
        private System.Windows.Forms.Button selectModelTable;
        private System.Windows.Forms.Button transform;
        private System.Windows.Forms.TextBox mapTextBox;
        private System.Windows.Forms.TextBox machineTextBox;
        private System.Windows.Forms.TextBox savaPathTextBox;
        private System.Windows.Forms.Button savePathButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label Time;
    }
}

