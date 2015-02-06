namespace K12.Sports.FitnessImportExport
{
     partial class InfoForm
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
               System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
               this.dataGridViewX1 = new DevComponents.DotNetBar.Controls.DataGridViewX();
               this.colClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
               this.colSeatNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
               this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
               this.colInfo = new System.Windows.Forms.DataGridViewTextBoxColumn();
               this.linkLabel1 = new System.Windows.Forms.LinkLabel();
               this.buttonX1 = new DevComponents.DotNetBar.ButtonX();
               this.linkLabel2 = new System.Windows.Forms.LinkLabel();
               ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).BeginInit();
               this.SuspendLayout();
               // 
               // dataGridViewX1
               // 
               this.dataGridViewX1.AllowUserToAddRows = false;
               this.dataGridViewX1.AllowUserToDeleteRows = false;
               this.dataGridViewX1.BackgroundColor = System.Drawing.Color.White;
               this.dataGridViewX1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
               this.dataGridViewX1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colClass,
            this.colSeatNo,
            this.colName,
            this.colInfo});
               dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
               dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
               dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
               dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
               dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
               dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.ControlText;
               dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
               this.dataGridViewX1.DefaultCellStyle = dataGridViewCellStyle3;
               this.dataGridViewX1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
               this.dataGridViewX1.Location = new System.Drawing.Point(10, 13);
               this.dataGridViewX1.Name = "dataGridViewX1";
               this.dataGridViewX1.RowTemplate.Height = 24;
               this.dataGridViewX1.Size = new System.Drawing.Size(632, 393);
               this.dataGridViewX1.TabIndex = 0;
               // 
               // colClass
               // 
               this.colClass.DataPropertyName = "_class";
               this.colClass.HeaderText = "班級";
               this.colClass.Name = "colClass";
               // 
               // colSeatNo
               // 
               this.colSeatNo.DataPropertyName = "_seatno";
               this.colSeatNo.HeaderText = "座號";
               this.colSeatNo.Name = "colSeatNo";
               // 
               // colName
               // 
               this.colName.DataPropertyName = "_name";
               this.colName.HeaderText = "姓名";
               this.colName.Name = "colName";
               // 
               // colInfo
               // 
               this.colInfo.DataPropertyName = "_info";
               this.colInfo.HeaderText = "資訊";
               this.colInfo.Name = "colInfo";
               this.colInfo.Width = 270;
               // 
               // linkLabel1
               // 
               this.linkLabel1.AutoSize = true;
               this.linkLabel1.BackColor = System.Drawing.Color.Transparent;
               this.linkLabel1.Location = new System.Drawing.Point(127, 426);
               this.linkLabel1.Name = "linkLabel1";
               this.linkLabel1.Size = new System.Drawing.Size(138, 17);
               this.linkLabel1.TabIndex = 1;
               this.linkLabel1.TabStop = true;
               this.linkLabel1.Text = "將清單學生加入待處理";
               this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
               // 
               // buttonX1
               // 
               this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
               this.buttonX1.AutoSize = true;
               this.buttonX1.BackColor = System.Drawing.Color.Transparent;
               this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
               this.buttonX1.Location = new System.Drawing.Point(559, 418);
               this.buttonX1.Name = "buttonX1";
               this.buttonX1.Size = new System.Drawing.Size(75, 25);
               this.buttonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
               this.buttonX1.TabIndex = 2;
               this.buttonX1.Text = "離開";
               this.buttonX1.Click += new System.EventHandler(this.buttonX1_Click);
               // 
               // linkLabel2
               // 
               this.linkLabel2.AutoSize = true;
               this.linkLabel2.BackColor = System.Drawing.Color.Transparent;
               this.linkLabel2.Location = new System.Drawing.Point(12, 426);
               this.linkLabel2.Name = "linkLabel2";
               this.linkLabel2.Size = new System.Drawing.Size(86, 17);
               this.linkLabel2.TabIndex = 3;
               this.linkLabel2.TabStop = true;
               this.linkLabel2.Text = "匯出訊息清單";
               this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
               // 
               // InfoForm
               // 
               this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
               this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
               this.ClientSize = new System.Drawing.Size(652, 454);
               this.Controls.Add(this.linkLabel2);
               this.Controls.Add(this.buttonX1);
               this.Controls.Add(this.linkLabel1);
               this.Controls.Add(this.dataGridViewX1);
               this.DoubleBuffered = true;
               this.Name = "InfoForm";
               this.Text = "常模轉換訊息";
               ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).EndInit();
               this.ResumeLayout(false);
               this.PerformLayout();

          }

          #endregion

          private DevComponents.DotNetBar.Controls.DataGridViewX dataGridViewX1;
          private System.Windows.Forms.LinkLabel linkLabel1;
          private DevComponents.DotNetBar.ButtonX buttonX1;
          private System.Windows.Forms.DataGridViewTextBoxColumn colClass;
          private System.Windows.Forms.DataGridViewTextBoxColumn colSeatNo;
          private System.Windows.Forms.DataGridViewTextBoxColumn colName;
          private System.Windows.Forms.DataGridViewTextBoxColumn colInfo;
          private System.Windows.Forms.LinkLabel linkLabel2;
     }
}