namespace AutoAttendance
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txt_ExcelPath1 = new System.Windows.Forms.TextBox();
            this.btn_InportExcel1 = new System.Windows.Forms.Button();
            this.btn_ExportReport = new System.Windows.Forms.Button();
            this.cob_BeginTime = new System.Windows.Forms.ComboBox();
            this.cob_EndTime = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_SaveSetting = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txt_SheetName2 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_SheetName1 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cob_Month = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txt_ExcelPath2 = new System.Windows.Forms.TextBox();
            this.btn_InportExcel2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // txt_ExcelPath1
            // 
            this.txt_ExcelPath1.Location = new System.Drawing.Point(8, 20);
            this.txt_ExcelPath1.Name = "txt_ExcelPath1";
            this.txt_ExcelPath1.ReadOnly = true;
            this.txt_ExcelPath1.Size = new System.Drawing.Size(306, 21);
            this.txt_ExcelPath1.TabIndex = 2;
            this.txt_ExcelPath1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btn_InportExcel1
            // 
            this.btn_InportExcel1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btn_InportExcel1.Location = new System.Drawing.Point(313, 19);
            this.btn_InportExcel1.Name = "btn_InportExcel1";
            this.btn_InportExcel1.Size = new System.Drawing.Size(24, 23);
            this.btn_InportExcel1.TabIndex = 4;
            this.btn_InportExcel1.Text = "十";
            this.btn_InportExcel1.UseVisualStyleBackColor = true;
            this.btn_InportExcel1.Click += new System.EventHandler(this.btn_InportExcel1_Click);
            // 
            // btn_ExportReport
            // 
            this.btn_ExportReport.BackColor = System.Drawing.Color.Lime;
            this.btn_ExportReport.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btn_ExportReport.Location = new System.Drawing.Point(214, 86);
            this.btn_ExportReport.Name = "btn_ExportReport";
            this.btn_ExportReport.Size = new System.Drawing.Size(84, 34);
            this.btn_ExportReport.TabIndex = 6;
            this.btn_ExportReport.Text = "导出报表";
            this.btn_ExportReport.UseVisualStyleBackColor = false;
            this.btn_ExportReport.Click += new System.EventHandler(this.btn_ExportReport_Click);
            // 
            // cob_BeginTime
            // 
            this.cob_BeginTime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cob_BeginTime.FormattingEnabled = true;
            this.cob_BeginTime.Items.AddRange(new object[] {
            "8:00",
            "8:30",
            "9:00",
            "16:00"});
            this.cob_BeginTime.Location = new System.Drawing.Point(77, 24);
            this.cob_BeginTime.Name = "cob_BeginTime";
            this.cob_BeginTime.Size = new System.Drawing.Size(73, 20);
            this.cob_BeginTime.TabIndex = 7;
            // 
            // cob_EndTime
            // 
            this.cob_EndTime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cob_EndTime.FormattingEnabled = true;
            this.cob_EndTime.Items.AddRange(new object[] {
            "16:00",
            "17:30",
            "18:00",
            "18:30",
            "00:00"});
            this.cob_EndTime.Location = new System.Drawing.Point(259, 24);
            this.cob_EndTime.Name = "cob_EndTime";
            this.cob_EndTime.Size = new System.Drawing.Size(73, 20);
            this.cob_EndTime.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "上班时间：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(188, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 10;
            this.label2.Text = "下班时间：";
            // 
            // btn_SaveSetting
            // 
            this.btn_SaveSetting.Location = new System.Drawing.Point(134, 125);
            this.btn_SaveSetting.Name = "btn_SaveSetting";
            this.btn_SaveSetting.Size = new System.Drawing.Size(74, 24);
            this.btn_SaveSetting.TabIndex = 11;
            this.btn_SaveSetting.Text = "保存设置";
            this.btn_SaveSetting.UseVisualStyleBackColor = true;
            this.btn_SaveSetting.Click += new System.EventHandler(this.btn_SaveSetting_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt_SheetName2);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txt_SheetName1);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btn_SaveSetting);
            this.groupBox1.Controls.Add(this.cob_BeginTime);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.cob_EndTime);
            this.groupBox1.Location = new System.Drawing.Point(11, 151);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(347, 155);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "系统设置：";
            // 
            // txt_SheetName2
            // 
            this.txt_SheetName2.Location = new System.Drawing.Point(119, 92);
            this.txt_SheetName2.Name = "txt_SheetName2";
            this.txt_SheetName2.Size = new System.Drawing.Size(213, 21);
            this.txt_SheetName2.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 95);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(107, 12);
            this.label6.TabIndex = 14;
            this.label6.Text = "九楼报表Sheet名：";
            // 
            // txt_SheetName1
            // 
            this.txt_SheetName1.Location = new System.Drawing.Point(119, 59);
            this.txt_SheetName1.Name = "txt_SheetName1";
            this.txt_SheetName1.Size = new System.Drawing.Size(213, 21);
            this.txt_SheetName1.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 62);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 12);
            this.label5.TabIndex = 12;
            this.label5.Text = "十楼报表Sheet名：";
            // 
            // cob_Month
            // 
            this.cob_Month.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cob_Month.FormattingEnabled = true;
            this.cob_Month.Location = new System.Drawing.Point(77, 94);
            this.cob_Month.Name = "cob_Month";
            this.cob_Month.Size = new System.Drawing.Size(73, 20);
            this.cob_Month.TabIndex = 13;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 14;
            this.label3.Text = "考勤月份：";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txt_ExcelPath2);
            this.groupBox2.Controls.Add(this.btn_InportExcel2);
            this.groupBox2.Controls.Add(this.txt_ExcelPath1);
            this.groupBox2.Controls.Add(this.btn_ExportReport);
            this.groupBox2.Controls.Add(this.btn_InportExcel1);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.cob_Month);
            this.groupBox2.Location = new System.Drawing.Point(11, 8);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(347, 130);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "考勤设置：";
            // 
            // txt_ExcelPath2
            // 
            this.txt_ExcelPath2.Location = new System.Drawing.Point(8, 47);
            this.txt_ExcelPath2.Name = "txt_ExcelPath2";
            this.txt_ExcelPath2.ReadOnly = true;
            this.txt_ExcelPath2.Size = new System.Drawing.Size(306, 21);
            this.txt_ExcelPath2.TabIndex = 15;
            this.txt_ExcelPath2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btn_InportExcel2
            // 
            this.btn_InportExcel2.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btn_InportExcel2.Location = new System.Drawing.Point(313, 46);
            this.btn_InportExcel2.Name = "btn_InportExcel2";
            this.btn_InportExcel2.Size = new System.Drawing.Size(24, 23);
            this.btn_InportExcel2.TabIndex = 16;
            this.btn_InportExcel2.Text = "九";
            this.btn_InportExcel2.UseVisualStyleBackColor = true;
            this.btn_InportExcel2.Click += new System.EventHandler(this.btn_InportExcel2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(366, 311);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "自动考勤工具";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TextBox txt_ExcelPath1;
        private System.Windows.Forms.Button btn_InportExcel1;
        private System.Windows.Forms.Button btn_ExportReport;
        private System.Windows.Forms.ComboBox cob_BeginTime;
        private System.Windows.Forms.ComboBox cob_EndTime;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_SaveSetting;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cob_Month;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_SheetName1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_SheetName2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txt_ExcelPath2;
        private System.Windows.Forms.Button btn_InportExcel2;
    }
}

