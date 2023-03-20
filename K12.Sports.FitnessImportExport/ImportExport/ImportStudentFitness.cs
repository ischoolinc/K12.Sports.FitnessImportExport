using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using Campus.Import2014;

namespace K12.Sports.FitnessImportExport.ImportExport
{
    class ImportStudentFitness : ImportWizard
    {
        private ImportOption _Option;
        
        // 新增
        private List<DAO.StudentFitnessRecord> _InsertRecList;
        // 更新
        private List<DAO.StudentFitnessRecord> _UpdateRecList;

        public ImportStudentFitness()
        {
            this.IsSplit = false;
            this.IsLog = false;
        }

        public override ImportAction GetSupportActions()
        {
            return ImportAction.InsertOrUpdate;
        }

        public override string GetValidateRule()
        {
            return Properties.Resources.ImportStudentFitnessRecordVal;
        }

        public override void Prepare(ImportOption Option)
        {
            
            _Option = Option;

            _InsertRecList = new List<DAO.StudentFitnessRecord>();
            _UpdateRecList = new List<DAO.StudentFitnessRecord>();
        }

        public override string Import(List<IRowStream> Rows)
        {
            _InsertRecList.Clear();
            _UpdateRecList.Clear();

            // 取得 Rows內學號, 學年度
            List<string> IDNumberList = new List<string>();
            Dictionary<string, string> SchoolYearDic = new Dictionary<string, string>();
            foreach (IRowStream row in Rows)
            {
                string IDNumber = Utility.GetIRowValueString(row, "身分證字號");
                string SchoolYear = Utility.GetIRowValueString(row, "學年度");
                
                if (string.IsNullOrEmpty(IDNumber)) continue;
                if (string.IsNullOrEmpty(SchoolYear)) continue;

                IDNumberList.Add(IDNumber);

                if (!SchoolYearDic.ContainsKey(SchoolYear))
                {
                    SchoolYearDic.Add(SchoolYear, "");
                }
            }

            // 透過身分證字號去取得學生ID
            Dictionary<string, string> IDNumDict = DAO.FDQuery.GetStudenByIDNumber(IDNumberList);
            
            // 根據學年度, 學生ID取得體適能的資料
            List<DAO.StudentFitnessRecord> fitnessRecList = DAO.StudentFitness.SelectByStudentIDListAndSchoolYear(IDNumDict.Values.ToList(), SchoolYearDic.Keys.ToList());
            Dictionary<string, List<DAO.StudentFitnessRecord>> fitnessRecDic = new Dictionary<string, List<DAO.StudentFitnessRecord>>();
            foreach (DAO.StudentFitnessRecord each in fitnessRecList)
            {
                if (!fitnessRecDic.ContainsKey(each.StudentID))
                {
                    fitnessRecDic.Add(each.StudentID, new List<DAO.StudentFitnessRecord>());
                }

                fitnessRecDic[each.StudentID].Add(each);
            }


            int totalCount = 0;
            // 判斷每一筆資料是要新增還是更新
            foreach(IRowStream row in Rows)
            {
                totalCount++;
                this.ImportProgress = totalCount;
                bool isInsert = true;   // 用來判斷此筆資料是否要新增
                DAO.StudentFitnessRecord fitnessRec = new DAO.StudentFitnessRecord();
                string IDNumber = Utility.GetIRowValueString(row, "身分證字號");
                int? SchoolYear = Utility.GetIRowValueInt(row, "學年度");
                DateTime? TestDate = Utility.GetIRowValueDateTime(row, "測驗日期");

                // 如果"學號"或"學年度"或"測驗日期"沒有資料, 換到下一筆
                if(string.IsNullOrEmpty(IDNumber)) continue;
                if(!SchoolYear.HasValue) continue;
                if(!TestDate.HasValue) continue;

                // 透過學號換成學生ID
                string StudentID = "";
                if(IDNumDict.ContainsKey(IDNumber))
                {
                    StudentID = IDNumDict[IDNumber];
                }
                else
                {
                    // 如果無法取得學生ID, 就換到下一筆
                    continue;
                }

                // 判斷此筆資料是否已在DB
                if (fitnessRecDic.ContainsKey(StudentID))
                {
                    foreach (DAO.StudentFitnessRecord rec in fitnessRecDic[StudentID])
                    {
                        if (rec.SchoolYear == SchoolYear)
                        {
                            isInsert = false;
                            fitnessRec = rec;
                            break;
                        }
                    }
                }
                else
                {
                    //沒有取得資料,不需動作
                }

                // 新增資料
                if(isInsert == true)
                {
                    #region 處理新增資料
                    // 學生系統編號
                    fitnessRec.StudentID = StudentID;
                    // 學年度
                    fitnessRec.SchoolYear = SchoolYear.Value;
                    // 測驗日期
                    fitnessRec.TestDate = TestDate.Value;
                    // 學校類別
                    fitnessRec.SchoolCategory = Utility.GetIRowValueString(row, "學校類別");
                    // 身高
                    fitnessRec.Height = Utility.GetIRowValueString(row, "身高");
                    // 身高常模
                    //fitnessRec.HeightDegree = Utility.GetIRowValueString(row, "身高常模");
                    // 體重
                    fitnessRec.Weight = Utility.GetIRowValueString(row, "體重");
                    // 體重常模
                    //fitnessRec.WeightDegree = Utility.GetIRowValueString(row, "體重常模");
                    // 坐姿體前彎
                    fitnessRec.SitAndReach = Utility.GetIRowValueString(row, "坐姿體前彎");
                    // 坐姿體前彎常模
                    fitnessRec.SitAndReachDegree = Utility.GetIRowValueString(row, "坐姿體前彎常模");
                    // 立定跳遠
                    fitnessRec.StandingLongJump = Utility.GetIRowValueString(row, "立定跳遠");
                    // 立定跳遠常模
                    fitnessRec.StandingLongJumpDegree = Utility.GetIRowValueString(row, "立定跳遠常模");
                    // 仰臥起坐
                    fitnessRec.SitUp = Utility.GetIRowValueString(row, "仰臥起坐");
                    // 仰臥起坐常模
                    fitnessRec.SitUpDegree = Utility.GetIRowValueString(row, "仰臥起坐常模");
                    // 心肺適能
                    fitnessRec.Cardiorespiratory = Utility.GetIRowValueString(row, "心肺適能");
                    // 心肺適能常模
                    fitnessRec.CardiorespiratoryDegree = Utility.GetIRowValueString(row, "心肺適能常模");
                    #endregion
                }
                // 更新資料
                else
                {
                    #region 處理更新資料
                    // "學號/座號", "學年度" 無法更新

                    // 測驗日期
                    if (_Option.SelectedFields.Contains("測驗日期"))
                        fitnessRec.TestDate = DateTime.Parse(Utility.GetIRowValueString(row, "測驗日期"));
                    // 學校類別
                    if (_Option.SelectedFields.Contains("學校類別"))
                        fitnessRec.SchoolCategory = Utility.GetIRowValueString(row, "學校類別");
                    // 身高
                    if (_Option.SelectedFields.Contains("身高"))
                        fitnessRec.Height = Utility.GetIRowValueString(row, "身高");
                    // 身高常模
                    //if (_Option.SelectedFields.Contains("身高常模"))
                    //    fitnessRec.HeightDegree = Utility.GetIRowValueString(row, "身高常模");
                    // 體重
                    if (_Option.SelectedFields.Contains("體重"))
                        fitnessRec.Weight = Utility.GetIRowValueString(row, "體重");
                    // 體重常模
                    //if (_Option.SelectedFields.Contains("體重常模"))
                    //    fitnessRec.WeightDegree = Utility.GetIRowValueString(row, "體重常模");
                    // 坐姿體前彎
                    if (_Option.SelectedFields.Contains("坐姿體前彎"))
                        fitnessRec.SitAndReach = Utility.GetIRowValueString(row, "坐姿體前彎");
                    // 坐姿體前彎常模
                    if (_Option.SelectedFields.Contains("坐姿體前彎常模"))
                        fitnessRec.SitAndReachDegree = Utility.GetIRowValueString(row, "坐姿體前彎常模");
                    // 立定跳遠
                    if (_Option.SelectedFields.Contains("立定跳遠"))
                        fitnessRec.StandingLongJump = Utility.GetIRowValueString(row, "立定跳遠");
                    // 立定跳遠常模
                    if (_Option.SelectedFields.Contains("立定跳遠常模"))
                        fitnessRec.StandingLongJumpDegree = Utility.GetIRowValueString(row, "立定跳遠常模");
                    // 仰臥起坐
                    if (_Option.SelectedFields.Contains("仰臥起坐"))
                        fitnessRec.SitUp = Utility.GetIRowValueString(row, "仰臥起坐");
                    // 仰臥起坐常模
                    if (_Option.SelectedFields.Contains("仰臥起坐常模"))
                        fitnessRec.SitUpDegree = Utility.GetIRowValueString(row, "仰臥起坐常模");
                    // 心肺適能
                    if (_Option.SelectedFields.Contains("心肺適能"))
                        fitnessRec.Cardiorespiratory = Utility.GetIRowValueString(row, "心肺適能");
                    // 心肺適能常模
                    if (_Option.SelectedFields.Contains("心肺適能常模"))
                        fitnessRec.CardiorespiratoryDegree = Utility.GetIRowValueString(row, "心肺適能常模");
                    #endregion
                }

                if(isInsert == true)
                    _InsertRecList.Add(fitnessRec);
                else
                    _UpdateRecList.Add(fitnessRec);
            }   // end of 判斷每一筆資料是要新增還是更新


            // 執行更新或新增
            if (_InsertRecList.Count >0 )
                DAO.StudentFitness.InsertByRecordList(_InsertRecList);
            if (_UpdateRecList.Count > 0)
                DAO.StudentFitness.UpdateByRecordList(_UpdateRecList);

            // Log
            Log.LogTransfer logTransfer = new Log.LogTransfer();
            StringBuilder logData = new StringBuilder();
            int insertCnt = _InsertRecList.Count; 
            int updateCnt = _UpdateRecList.Count;
            logData.Append("總共匯入").Append((insertCnt + updateCnt)).Append("筆,");
            logData.Append("新增:").Append(insertCnt).Append("筆,");
            logData.Append("更新:").Append(updateCnt).Append("筆");
            logTransfer.SaveLog("體適能", "匯入", "student", "", logData);

            return "";
        }
    }
}
