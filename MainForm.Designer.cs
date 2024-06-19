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
            this.selectMappingTable = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(344, 33);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(352, 46);
            this.label1.TabIndex = 0;
            this.label1.Text = "三厂Spec測試軟件";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // selectModelTable
            // 
            this.selectModelTable.BackColor = System.Drawing.Color.Salmon;
            this.selectModelTable.Location = new System.Drawing.Point(133, 161);
            this.selectModelTable.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.selectModelTable.Name = "selectModelTable";
            this.selectModelTable.Size = new System.Drawing.Size(110, 35);
            this.selectModelTable.TabIndex = 2;
            this.selectModelTable.Text = "選擇Spec文件";
            this.selectModelTable.UseVisualStyleBackColor = false;
            this.selectModelTable.Click += new System.EventHandler(this.selectModelTable_Click);
            // 
            // transform
            // 
            this.transform.BackColor = System.Drawing.Color.Orange;
            this.transform.ForeColor = System.Drawing.SystemColors.Desktop;
            this.transform.Location = new System.Drawing.Point(210, 275);
            this.transform.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.transform.Name = "transform";
            this.transform.Size = new System.Drawing.Size(188, 35);
            this.transform.TabIndex = 3;
            this.transform.Text = "轉檔";
            this.transform.UseVisualStyleBackColor = false;
            this.transform.Click += new System.EventHandler(this.transform_click);
            // 
            // mapTextBox
            // 
            this.mapTextBox.Location = new System.Drawing.Point(276, 117);
            this.mapTextBox.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.mapTextBox.MaximumSize = new System.Drawing.Size(335, 500);
            this.mapTextBox.Name = "mapTextBox";
            this.mapTextBox.Size = new System.Drawing.Size(246, 21);
            this.mapTextBox.TabIndex = 4;
            this.mapTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // machineTextBox
            // 
            this.machineTextBox.Location = new System.Drawing.Point(276, 169);
            this.machineTextBox.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.machineTextBox.Name = "machineTextBox";
            this.machineTextBox.Size = new System.Drawing.Size(246, 21);
            this.machineTextBox.TabIndex = 6;
            this.machineTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // savaPathTextBox
            // 
            this.savaPathTextBox.Location = new System.Drawing.Point(276, 220);
            this.savaPathTextBox.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.savaPathTextBox.Name = "savaPathTextBox";
            this.savaPathTextBox.Size = new System.Drawing.Size(246, 21);
            this.savaPathTextBox.TabIndex = 7;
            this.savaPathTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // savePathButton
            // 
            this.savePathButton.BackColor = System.Drawing.Color.Salmon;
            this.savePathButton.Location = new System.Drawing.Point(133, 212);
            this.savePathButton.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.savePathButton.Name = "savePathButton";
            this.savePathButton.Size = new System.Drawing.Size(110, 35);
            this.savePathButton.TabIndex = 8;
            this.savePathButton.Text = "選擇存放路徑";
            this.savePathButton.UseVisualStyleBackColor = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(92, 67);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 37);
            this.label2.TabIndex = 9;
            this.label2.Text = "路徑";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(39, 23);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(144, 40);
            this.pictureBox1.TabIndex = 31;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(384, 331);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 12);
            this.label3.TabIndex = 32;
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // Time
            // 
            this.Time.AutoSize = true;
            this.Time.Location = new System.Drawing.Point(411, 331);
            this.Time.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Time.Name = "Time";
            this.Time.Size = new System.Drawing.Size(0, 12);
            this.Time.TabIndex = 33;
            // 
            // selectMappingTable
            // 
            this.selectMappingTable.BackColor = System.Drawing.Color.Salmon;
            this.selectMappingTable.ForeColor = System.Drawing.SystemColors.Desktop;
            this.selectMappingTable.Location = new System.Drawing.Point(133, 110);
            this.selectMappingTable.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.selectMappingTable.Name = "selectMappingTable";
            this.selectMappingTable.Size = new System.Drawing.Size(110, 32);
            this.selectMappingTable.TabIndex = 1;
            this.selectMappingTable.Text = "映射表路径";
            this.selectMappingTable.UseVisualStyleBackColor = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 360);
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
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
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
        private System.Windows.Forms.Button selectMappingTable;
    }
}

