using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;

namespace K12.Sports.FitnessImportExport.Forms
{
    public partial class FrmFitnessRecord : BaseForm
    {
        public enum accessType { Insert, Edit }
        StudentRecord _studRec;
        DAO.StudentFitnessRecord _fitnessRec;
        accessType _actType;
        Log.LogTransfer _LogTransfer;
        readonly string _FrmTitleAdd = "體適能資料新增";
        readonly string _FrmTitleEdit = "體適能資料修改";


        public FrmFitnessRecord(DAO.StudentFitnessRecord rec, accessType actType)
        {
            InitializeComponent();

            _studRec = Student.SelectByID(rec.StudentID);
            _fitnessRec = rec;
            _actType = actType;
            _LogTransfer = new Log.LogTransfer();

            if (_actType == accessType.Edit)
            {
                this.Text = _FrmTitleEdit;
                //修改模式無法變更學年度
                this.integerInput1.Enabled = false;
            }
            else
            {
                this.Text = _FrmTitleAdd;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmFitnessRecord_Load(object sender, EventArgs e)
        {
            LoadDefaultDataToForm();
            LoadUpdateDataToForm();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFormDataToFitnessRec();

            Utility.SetLogData(_LogTransfer, _fitnessRec);

            string studStr = "學號:" + _studRec.StudentNumber + ",姓名:" + _studRec.Name + ",";
            if (_actType == accessType.Insert)
            {
                // 檢查是否有重複 (SchoolYear+studentID)
                bool isPass = true;
                List<DAO.StudentFitnessRecord> recList = DAO.StudentFitness.SelectByStudentIDAndSchoolYear(_fitnessRec.StudentID, _fitnessRec.SchoolYear);

                foreach (DAO.StudentFitnessRecord rec in recList)
                {
                    if (rec.SchoolYear == _fitnessRec.SchoolYear)
                        isPass = false;

                    //if(rec.TestDate == _fitnessRec.TestDate)
                    //    isPass = false;
                }

                if (isPass == true)
                {
                    _LogTransfer.SaveInsertLog("學生.體適能-新增", "新增", studStr, "", "student", _studRec.ID);
                    // insert data
                    DAO.StudentFitness.InsertByRecord(_fitnessRec);
                }
                else
                {
                    // data duplicate
                    FISCA.Presentation.Controls.MsgBox.Show("已有相同的學年度，無法新增。");
                    return;
                }
            }
            else
            {
                _LogTransfer.SaveChangeLog("學生.體適能-修改", "修改", studStr, "", "student", _studRec.ID);
                // update
                DAO.StudentFitness.UpdateByRecord(_fitnessRec);
            }

            FISCA.Presentation.Controls.MsgBox.Show("儲存完成。");
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void LoadDefaultDataToForm()
        {
            lbStudentName.Text = _studRec.Name;

            if (_studRec.Class != null)
            {
                if (_studRec.Class.GradeYear.HasValue)
                    lbClassGradeYear.Text = _studRec.Class.GradeYear.Value.ToString();
                lbClassName.Text = _studRec.Class.Name;
            }
        }

        private void LoadUpdateDataToForm()
        {
            integerInput1.Value = _fitnessRec.SchoolYear;
            dateTimeInput1.Value = _fitnessRec.TestDate;

            txtSchoolCategory.Text = _fitnessRec.SchoolCategory;

            txtHeight.Text = _fitnessRec.Height;
            txtHeightDegree.Text = _fitnessRec.HeightDegree;

            txtWeight.Text = _fitnessRec.Weight;
            txtWeightDegree.Text = _fitnessRec.WeightDegree;

            txtSitAndReach.Text = _fitnessRec.SitAndReach;
            txtSitAndReachDegree.Text = _fitnessRec.SitAndReachDegree;

            txtStandingLongJump.Text = _fitnessRec.StandingLongJump;
            txtStandingLongJumpDegree.Text = _fitnessRec.StandingLongJumpDegree;

            txtSitUp.Text = _fitnessRec.SitUp;
            txtSitUpDegree.Text = _fitnessRec.SitUpDegree;

            txtCardiorespiratory.Text = _fitnessRec.Cardiorespiratory;
            txtCardiorespiratoryDegree.Text = _fitnessRec.CardiorespiratoryDegree;

            _LogTransfer.Clear();
            Utility.SetLogData(_LogTransfer, _fitnessRec);
        }

        private void SaveFormDataToFitnessRec()
        {
            _fitnessRec.SchoolYear = integerInput1.Value;
            _fitnessRec.TestDate = dateTimeInput1.Value;

            _fitnessRec.SchoolCategory = txtSchoolCategory.Text;

            _fitnessRec.Height = txtHeight.Text;
            _fitnessRec.HeightDegree = txtHeightDegree.Text;

            _fitnessRec.Weight = txtWeight.Text;
            _fitnessRec.WeightDegree = txtWeightDegree.Text;

            _fitnessRec.SitAndReach = txtSitAndReach.Text;
            _fitnessRec.SitAndReachDegree = txtSitAndReachDegree.Text;

            _fitnessRec.StandingLongJump = txtStandingLongJump.Text;
            _fitnessRec.StandingLongJumpDegree = txtStandingLongJumpDegree.Text;

            _fitnessRec.SitUp = txtSitUp.Text;
            _fitnessRec.SitUpDegree = txtSitUpDegree.Text;

            _fitnessRec.Cardiorespiratory = txtCardiorespiratory.Text;
            _fitnessRec.CardiorespiratoryDegree = txtCardiorespiratoryDegree.Text;
        }
    }
}
