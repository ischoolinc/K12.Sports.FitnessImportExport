namespace K12.Sports.FitnessImportExport
{
     partial class ComparisonForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ComparisonForm));
            this.lbHelp1 = new DevComponents.DotNetBar.LabelX();
            this.intSchoolYear = new DevComponents.Editors.IntegerInput();
            this.btnStart = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.cbCover = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.lbHelp2 = new DevComponents.DotNetBar.LabelX();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.intSchoolYear)).BeginInit();
            this.SuspendLayout();
            // 
            // lbHelp1
            // 
            this.lbHelp1.AutoSize = true;
            this.lbHelp1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbHelp1.BackgroundStyle.Class = "";
            this.lbHelp1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp1.Location = new System.Drawing.Point(57, 32);
            this.lbHelp1.Name = "lbHelp1";
            this.lbHelp1.Size = new System.Drawing.Size(87, 21);
            this.lbHelp1.TabIndex = 0;
            this.lbHelp1.Text = "請選擇學年度";
            // 
            // intSchoolYear
            // 
            this.intSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.intSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.intSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.intSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.intSchoolYear.Location = new System.Drawing.Point(164, 30);
            this.intSchoolYear.MaxValue = 999;
            this.intSchoolYear.MinValue = 90;
            this.intSchoolYear.Name = "intSchoolYear";
            this.intSchoolYear.ShowUpDown = true;
            this.intSchoolYear.Size = new System.Drawing.Size(84, 25);
            this.intSchoolYear.TabIndex = 1;
            this.intSchoolYear.Value = 90;
            // 
            // btnStart
            // 
            this.btnStart.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStart.AutoSize = true;
            this.btnStart.BackColor = System.Drawing.Color.Transparent;
            this.btnStart.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnStart.Location = new System.Drawing.Point(355, 561);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 25);
            this.btnStart.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "常模換算";
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(437, 561);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 3;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // cbCover
            // 
            this.cbCover.AutoSize = true;
            this.cbCover.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbCover.BackgroundStyle.Class = "";
            this.cbCover.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbCover.Location = new System.Drawing.Point(276, 34);
            this.cbCover.Name = "cbCover";
            this.cbCover.Size = new System.Drawing.Size(174, 21);
            this.cbCover.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbCover.TabIndex = 4;
            this.cbCover.Text = "常模欄位有值時進行覆蓋";
            // 
            // lbHelp2
            // 
            this.lbHelp2.AutoSize = true;
            this.lbHelp2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbHelp2.BackgroundStyle.Class = "";
            this.lbHelp2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp2.Location = new System.Drawing.Point(22, 83);
            this.lbHelp2.Name = "lbHelp2";
            this.lbHelp2.Size = new System.Drawing.Size(476, 194);
            this.lbHelp2.TabIndex = 5;
            this.lbHelp2.Text = resources.GetString("lbHelp2.Text");
            // 
            // linkLabel2
            // 
            this.linkLabel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.BackColor = System.Drawing.Color.Transparent;
            this.linkLabel2.Location = new System.Drawing.Point(19, 569);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(73, 17);
            this.linkLabel2.TabIndex = 7;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "常模對照表";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(22, 311);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(502, 194);
            this.labelX1.TabIndex = 8;
            this.labelX1.Text = resources.GetString("labelX1.Text");
            // 
            // ComparisonForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 595);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.lbHelp2);
            this.Controls.Add(this.cbCover);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.intSchoolYear);
            this.Controls.Add(this.lbHelp1);
            this.DoubleBuffered = true;
            this.Name = "ComparisonForm";
            this.Text = "體適能常模轉換";
            this.Load += new System.EventHandler(this.ComparisonForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.intSchoolYear)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

          }

          #endregion

          private DevComponents.DotNetBar.LabelX lbHelp1;
          private DevComponents.Editors.IntegerInput intSchoolYear;
          private DevComponents.DotNetBar.ButtonX btnStart;
          private DevComponents.DotNetBar.ButtonX btnExit;
          private DevComponents.DotNetBar.Controls.CheckBoxX cbCover;
          private DevComponents.DotNetBar.LabelX lbHelp2;
          private System.Windows.Forms.LinkLabel linkLabel2;
          private DevComponents.DotNetBar.LabelX labelX1;
     }
}