using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Permission;
using Campus.Windows;
using FISCA.Presentation;

namespace K12.Sports.FitnessImportExport.DetailContents
{
    [FeatureCode(Permissions.KeyFitnessContent, "體適能")]
    public partial class StudentFitnessContent : DetailContent
    {
        // 背景處理
        private BackgroundWorker _bgWorker;
        bool _isBusy = false;
        List<DAO.StudentFitnessRecord> _FitnessRecList = new List<DAO.StudentFitnessRecord>();

        string EventCode = "Fitness.now.Comparison";

        public StudentFitnessContent()
        {
            InitializeComponent();

            // 資料項目名稱
            this.Group = Global._ModuleName;

            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);

            // 判斷權限
            if(Permissions.IsEditableFitnessContent == true)
            {
                this.btnDelete.Enabled = true;
                this.btnInsert.Enabled = true;
                this.btnUpdate.Enabled = true;
            }
            else
            {
                this.btnDelete.Enabled = false;
                this.btnInsert.Enabled = false;
                this.btnUpdate.Enabled = false;
            }

            FISCA.InteractionService.SubscribeEvent(EventCode, (sender, args) =>
            {
                 this.Loading = true;
                 _BGRun();
            });
        }

        // 切換學生時
        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            this.Loading = true;

            _BGRun();
        }
        
        // 把背景執行緒叫起來做事
        private void _BGRun()
        {
            if (_bgWorker.IsBusy)
                _isBusy = true;
            else
                _bgWorker.RunWorkerAsync();
        }

        // 開始背景處理
        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // 取得資料
            _FitnessRecList = DAO.StudentFitness.SelectByStudentID(this.PrimaryKey).OrderByDescending(x => x.SchoolYear).ThenByDescending(x => x.TestDate).ToList();
        }

        // 假如背景處理結束, 就會呼叫本method
        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }

            // 載入資料至畫面
            LoadDataToView();
        }

        private void LoadDataToView()
        {
            this.Loading = false;
            lvFitnessList.Items.Clear();
            ListViewItem item = null;
            foreach(DAO.StudentFitnessRecord rec in _FitnessRecList)
            {
                item = ConvertToListViewItem(rec);
                
                lvFitnessList.Items.Add(item);
                
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            DAO.StudentFitnessRecord rec = new DAO.StudentFitnessRecord();
            rec.StudentID = PrimaryKey;
            rec.SchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            rec.TestDate = DateTime.Today;
            Forms.FrmFitnessRecord frm = new Forms.FrmFitnessRecord(rec, Forms.FrmFitnessRecord.accessType.Insert);
            if(frm.ShowDialog() == DialogResult.OK)
                _BGRun();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (lvFitnessList.SelectedItems.Count > 0)
            {
                DAO.StudentFitnessRecord rec = lvFitnessList.SelectedItems[0].Tag as DAO.StudentFitnessRecord;
                if(rec != null)
                {
                    Forms.FrmFitnessRecord frm = new Forms.FrmFitnessRecord(rec, Forms.FrmFitnessRecord.accessType.Edit);
                    if(frm.ShowDialog() == DialogResult.OK)
                        _BGRun();
                }

            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
            }

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if(lvFitnessList.SelectedItems.Count > 0)
            {
                DAO.StudentFitnessRecord rec = lvFitnessList.SelectedItems[0].Tag as DAO.StudentFitnessRecord;
                if (FISCA.Presentation.Controls.MsgBox.Show("請問是否確定是刪除體適能紀錄?", "刪除體適能紀錄", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    K12.Data.StudentRecord studRec = K12.Data.Student.SelectByID(PrimaryKey);
                    string studStr = "學號:" + studRec.StudentNumber + ",姓名:" + studRec.Name + ",";

                    Log.LogTransfer LogTransfer = new Log.LogTransfer();
                    Utility.SetLogData(LogTransfer, rec);
                    LogTransfer.SaveDeleteLog("學生.體適能-刪除", "刪除", studStr, "", "student", studRec.ID);
                    
                    DAO.StudentFitness.DeleteByRecord(rec);

                    _BGRun();
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
            }
        }
        
        private ListViewItem ConvertToListViewItem(DAO.StudentFitnessRecord rec)
        {
            ListViewItem item = new ListViewItem();

            item.Tag = rec;
            // 學年度
            item.Text = rec.SchoolYear.ToString();
            // 測驗日期
            item.SubItems.Add(rec.TestDate.ToShortDateString());
            // 學校類別
            item.SubItems.Add(rec.SchoolCategory);
            // 身高
            item.SubItems.Add(rec.Height);
            // 身高常模
            //item.SubItems.Add(rec.HeightDegree);
            // 體重
            item.SubItems.Add(rec.Weight);
            // 體重常模
            //item.SubItems.Add(rec.WeightDegree);
            // 坐姿體前彎
            item.SubItems.Add(rec.SitAndReach);
            // 坐姿體前彎常模
            item.SubItems.Add(rec.SitAndReachDegree);
            // 立定跳遠
            item.SubItems.Add(rec.StandingLongJump);
            // 立定跳遠常模
            item.SubItems.Add(rec.StandingLongJumpDegree);
            // 仰臥起坐
            item.SubItems.Add(rec.SitUp);
            // 仰臥起坐常模
            item.SubItems.Add(rec.SitUpDegree);
            // 心肺適能(心肺耐力)
            item.SubItems.Add(rec.Cardiorespiratory);
            // 心肺適能常模(心肺耐力)
            item.SubItems.Add(rec.CardiorespiratoryDegree);
            //  仰臥捲腹
            item.SubItems.Add(rec.Curl);
            //  仰臥捲腹常模
            item.SubItems.Add(rec.CurlDegree);
            //  漸速耐力跑
            item.SubItems.Add(rec.Pacer);
            //  漸速耐力跑常模
            item.SubItems.Add(rec.PacerDegree);

            return item;
        }
    }
}
