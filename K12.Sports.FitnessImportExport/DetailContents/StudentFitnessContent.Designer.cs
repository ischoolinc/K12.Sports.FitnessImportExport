namespace K12.Sports.FitnessImportExport.DetailContents
{
    partial class StudentFitnessContent
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
               this.lvFitnessList = new DevComponents.DotNetBar.Controls.ListViewEx();
               this.colSemester = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colTestDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colSchoolCategory = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colHeight = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colHeightDegree = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colWeight = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colWeightDegree = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colSitAndReach = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colSitAndReachDegree = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colStandingLongJump = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colStandingLongJumpDegree = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colSitUp = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colSitUpDegree = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colCardiorespiratory = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.colCardiorespiratoryDegree = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
               this.btnInsert = new DevComponents.DotNetBar.ButtonX();
               this.btnUpdate = new DevComponents.DotNetBar.ButtonX();
               this.btnDelete = new DevComponents.DotNetBar.ButtonX();
               this.SuspendLayout();
               // 
               // lvFitnessList
               // 
               // 
               // 
               // 
               this.lvFitnessList.Border.Class = "ListViewBorder";
               this.lvFitnessList.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
               this.lvFitnessList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colSemester,
            this.colTestDate,
            this.colSchoolCategory,
            this.colHeight,
            this.colHeightDegree,
            this.colWeight,
            this.colWeightDegree,
            this.colSitAndReach,
            this.colSitAndReachDegree,
            this.colStandingLongJump,
            this.colStandingLongJumpDegree,
            this.colSitUp,
            this.colSitUpDegree,
            this.colCardiorespiratory,
            this.colCardiorespiratoryDegree});
               this.lvFitnessList.FullRowSelect = true;
               this.lvFitnessList.Location = new System.Drawing.Point(13, 8);
               this.lvFitnessList.Name = "lvFitnessList";
               this.lvFitnessList.Size = new System.Drawing.Size(524, 120);
               this.lvFitnessList.TabIndex = 0;
               this.lvFitnessList.UseCompatibleStateImageBehavior = false;
               this.lvFitnessList.View = System.Windows.Forms.View.Details;
               // 
               // colSemester
               // 
               this.colSemester.Text = "學年度";
               // 
               // colTestDate
               // 
               this.colTestDate.Text = "測驗日期";
               this.colTestDate.Width = 75;
               // 
               // colSchoolCategory
               // 
               this.colSchoolCategory.Text = "學校類別";
               this.colSchoolCategory.Width = 75;
               // 
               // colHeight
               // 
               this.colHeight.Text = "身高";
               // 
               // colHeightDegree
               // 
               this.colHeightDegree.Text = "身高常模";
               this.colHeightDegree.Width = 75;
               // 
               // colWeight
               // 
               this.colWeight.Text = "體重";
               // 
               // colWeightDegree
               // 
               this.colWeightDegree.Text = "體重常模";
               this.colWeightDegree.Width = 75;
               // 
               // colSitAndReach
               // 
               this.colSitAndReach.Text = "坐姿體前彎";
               this.colSitAndReach.Width = 90;
               // 
               // colSitAndReachDegree
               // 
               this.colSitAndReachDegree.Text = "坐姿體前彎常模";
               this.colSitAndReachDegree.Width = 120;
               // 
               // colStandingLongJump
               // 
               this.colStandingLongJump.Text = "立定跳遠";
               this.colStandingLongJump.Width = 75;
               // 
               // colStandingLongJumpDegree
               // 
               this.colStandingLongJumpDegree.Text = "立定跳遠常模";
               this.colStandingLongJumpDegree.Width = 100;
               // 
               // colSitUp
               // 
               this.colSitUp.Text = "仰臥起坐";
               this.colSitUp.Width = 75;
               // 
               // colSitUpDegree
               // 
               this.colSitUpDegree.Text = "仰臥起坐常模";
               this.colSitUpDegree.Width = 100;
               // 
               // colCardiorespiratory
               // 
               this.colCardiorespiratory.Text = "心肺適能";
               this.colCardiorespiratory.Width = 75;
               // 
               // colCardiorespiratoryDegree
               // 
               this.colCardiorespiratoryDegree.Text = "心肺適能常模";
               this.colCardiorespiratoryDegree.Width = 100;
               // 
               // btnInsert
               // 
               this.btnInsert.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
               this.btnInsert.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
               this.btnInsert.Location = new System.Drawing.Point(13, 134);
               this.btnInsert.Name = "btnInsert";
               this.btnInsert.Size = new System.Drawing.Size(75, 23);
               this.btnInsert.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
               this.btnInsert.TabIndex = 1;
               this.btnInsert.Text = "新增";
               this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
               // 
               // btnUpdate
               // 
               this.btnUpdate.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
               this.btnUpdate.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
               this.btnUpdate.Location = new System.Drawing.Point(94, 134);
               this.btnUpdate.Name = "btnUpdate";
               this.btnUpdate.Size = new System.Drawing.Size(75, 23);
               this.btnUpdate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
               this.btnUpdate.TabIndex = 2;
               this.btnUpdate.Text = "修改";
               this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
               // 
               // btnDelete
               // 
               this.btnDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
               this.btnDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
               this.btnDelete.Location = new System.Drawing.Point(175, 134);
               this.btnDelete.Name = "btnDelete";
               this.btnDelete.Size = new System.Drawing.Size(75, 23);
               this.btnDelete.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
               this.btnDelete.TabIndex = 3;
               this.btnDelete.Text = "刪除";
               this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
               // 
               // StudentFitnessContent
               // 
               this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
               this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
               this.Controls.Add(this.btnDelete);
               this.Controls.Add(this.btnUpdate);
               this.Controls.Add(this.btnInsert);
               this.Controls.Add(this.lvFitnessList);
               this.Name = "StudentFitnessContent";
               this.Size = new System.Drawing.Size(550, 175);
               this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ListViewEx lvFitnessList;
        private System.Windows.Forms.ColumnHeader colSemester;
        private System.Windows.Forms.ColumnHeader colTestDate;
        private System.Windows.Forms.ColumnHeader colHeight;
        private DevComponents.DotNetBar.ButtonX btnInsert;
        private DevComponents.DotNetBar.ButtonX btnUpdate;
        private DevComponents.DotNetBar.ButtonX btnDelete;
        private System.Windows.Forms.ColumnHeader colHeightDegree;
        private System.Windows.Forms.ColumnHeader colWeight;
        private System.Windows.Forms.ColumnHeader colWeightDegree;
        private System.Windows.Forms.ColumnHeader colSitAndReach;
        private System.Windows.Forms.ColumnHeader colSitAndReachDegree;
        private System.Windows.Forms.ColumnHeader colStandingLongJump;
        private System.Windows.Forms.ColumnHeader colStandingLongJumpDegree;
        private System.Windows.Forms.ColumnHeader colSitUp;
        private System.Windows.Forms.ColumnHeader colSitUpDegree;
        private System.Windows.Forms.ColumnHeader colCardiorespiratory;
        private System.Windows.Forms.ColumnHeader colCardiorespiratoryDegree;
        private System.Windows.Forms.ColumnHeader colSchoolCategory;
    }
}
