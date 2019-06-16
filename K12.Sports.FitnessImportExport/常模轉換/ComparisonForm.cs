using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Sports.FitnessImportExport.DAO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace K12.Sports.FitnessImportExport
{
    public partial class ComparisonForm : BaseForm
    {
        BackgroundWorker BGW = new BackgroundWorker();

        //事件物件
        EventHandler eh;
        //事件代碼
        string EventCode = "Fitness.now.Comparison";

        //全部覆蓋
        bool CheckCover = false;
        //擇優覆蓋
        bool CheckPreferred = false;

        List<FitInfo> FitInfoList { get; set; }

        /// <summary>
        /// 常模大小對照用
        /// </summary>
        Dictionary<string, int> FitnessDic = new Dictionary<string, int>();

        public ComparisonForm()
        {
            InitializeComponent();
        }

        private void ComparisonForm_Load(object sender, EventArgs e)
        {
            BGW.RunWorkerCompleted += BGW_RunWorkerCompleted;
            BGW.DoWork += BGW_DoWork;

            FitnessDic.Add("請加強", 1);
            FitnessDic.Add("中等", 2);
            FitnessDic.Add("銅牌", 3);
            FitnessDic.Add("銀牌", 4);
            FitnessDic.Add("金牌", 5);

            //透過代碼,取得事件引發器
            eh = FISCA.InteractionService.PublishEvent(EventCode);

            intSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!BGW.IsBusy)
            {
                CheckPreferred = checkBoxX1.Checked;
                CheckCover = checkBoxX2.Checked;

                btnStart.Enabled = false;
                BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("系統忙碌中,稍後再試!");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得學生資料
            Dictionary<string, KeyBoStudent> _student = tool.GetStudentList(K12.Presentation.NLDPanels.Student.SelectedSource);

            FitInfoList = new List<FitInfo>();

            //取得體適能資料
            List<StudentFitnessRecord> SFitnessList = tool._A.Select<StudentFitnessRecord>(string.Format("ref_student_id in ('{0}') and school_year={1}", string.Join("','", K12.Presentation.NLDPanels.Student.SelectedSource), intSchoolYear.Value.ToString()));
            foreach (StudentFitnessRecord each in SFitnessList)
            {
                #region 每一筆資料

                if (_student.ContainsKey(each.StudentID))
                {
                    KeyBoStudent student = _student[each.StudentID];
                    student.sfr = each; //記錄LOG
                    student.sfr_Log = each.CopyExtension(); //記錄LOG

                    #region 說明

                    //換算實際監測日期的年齡
                    //(一)體適能檢測之年齡計算方式，以7個月為界，
                    //檢測年月與出生年月相減，所得月分達7個月以上，則進升一歲。
                    //舉例如下：
                    //1.民國88年3月出生，於民國101年10月進行檢測，
                    //年齡計算為101(年)-88(年)=13；10(月)-3(月)=7，因達7個月，故進升1歲，其年齡為14歲。

                    //2.民國88年5月出生，於民國101年10月進行檢測，
                    //年齡計算為101(年)-88(年)=13；10(月)-5(月)=5，因未達7個月，故其年齡為13歲。

                    //(二)受測學生超過16歲者，以16歲之門檻標準計分
                    //(2015/2/26 new)
                    //3.未滿13歲者均以13歲之常模進行鑑測(103+學年度之資料)

                    #endregion


                    //DateTime dt1 = new DateTime(1999, 11, 2); //生日
                    //DateTime dt2 = new DateTime(2014, 12, 12); //測驗

                    //student.Birthdate = dt1;
                    //each.TestDate = dt2;

                    if (each.TestDate != null)
                    {
                        if (student.Birthdate.HasValue)
                        {
                            if (each.TestDate > student.Birthdate)
                            {
                                #region 邏輯核心

                                student.Age = getAge2(student.Birthdate.Value, each.TestDate);

                                //當資料學年度,是大於102時
                                //採用13足歲之規則
                                if (each.SchoolYear >= 103)
                                {
                                    //new - 凡未滿13歲均以13歲為基準進行換算
                                    if (student.Age.年 < 13)
                                    {
                                        student.Age.鑑測年齡 = 13;
                                    }
                                    else
                                    {
                                        if (student.Age.月 >= 7)
                                        {
                                            //大於7個月則年齡加一
                                            student.Age.鑑測年齡++;
                                        }
                                    }
                                }
                                else
                                {
                                    if (student.Age.月 >= 7)
                                    {
                                        //大於7個月則年齡加一
                                        student.Age.鑑測年齡++;
                                    }
                                }
                            }
                            else
                            {
                                //沒有年齡,表示該體適能資料無法換算
                                FitInfo info = new FitInfo(student);
                                info._info = "測驗日期不可能在出生之前!";
                                FitInfoList.Add(info);
                            }

                            #endregion
                        }
                        else
                        {
                            //沒有年齡,表示該體適能資料無法換算
                            FitInfo info = new FitInfo(student);
                            info._info = "沒有生日!";
                            FitInfoList.Add(info);
                        }
                    }
                    else
                    {
                        //沒有測驗日期
                        //沒有年齡,表示該體適能資料無法換算
                        FitInfo info = new FitInfo(student);
                        info._info = "未輸入體適能測驗日期!!";
                        FitInfoList.Add(info);
                    }
                }

                #endregion
            }

            List<StudentFitnessRecord> UpdateList = new List<StudentFitnessRecord>();


            //取得對照表,整理為可對照清單
            XmlElement xml = DSXmlHelper.LoadXml(Properties.Resources.Sports_Fitness_Comparison);
            SuperComparison _sc = new SuperComparison(xml);

            foreach (StudentFitnessRecord sfr in SFitnessList)
            {
                //包含這位學生
                if (_student.ContainsKey(sfr.StudentID))
                {
                    KeyBoStudent student = _student[sfr.StudentID];

                    if (student.Gender == "男")
                    {
                        if (CheckPreferred) //2019/6/14 - 擇優覆蓋
                        {
                            //資料是否存在
                            string SitUpDegreeA = sfr.SitUpDegree.Trim();
                            string SitUpDegreeB = GetMeValue(sfr.SitUp.Trim(), "仰臥起坐", student, _sc.Boy_仰臥起坐);
                            sfr.SitUpDegree = CompareFitness(SitUpDegreeA, SitUpDegreeB);

                            string SitAndReachDegreeA = sfr.SitAndReachDegree.Trim();
                            string SitAndReachDegreeB = GetMeValue(sfr.SitAndReach.Trim(), "坐姿體前彎", student, _sc.Boy_坐姿體前彎);
                            sfr.SitAndReachDegree = CompareFitness(SitAndReachDegreeA, SitAndReachDegreeB);

                            string StandingLongJumpDegreeA = sfr.StandingLongJumpDegree.Trim();
                            string StandingLongJumpDegreeB = GetMeValue(sfr.StandingLongJump.Trim(), "立定跳遠", student, _sc.Boy_立定跳遠);
                            sfr.StandingLongJumpDegree = CompareFitness(StandingLongJumpDegreeA, StandingLongJumpDegreeB);

                            string CardiorespiratoryDegreeA = sfr.CardiorespiratoryDegree.Trim();
                            string CardiorespiratoryDegreeB = GetMeValue(dotorsec(sfr.Cardiorespiratory.Trim()), "心肺適能", student, _sc.Boy_心肺適能);
                            sfr.CardiorespiratoryDegree = CompareFitness(CardiorespiratoryDegreeA, CardiorespiratoryDegreeB);
                        }
                        else if (CheckCover)
                        {
                            //全部覆蓋
                            sfr.SitUpDegree = GetMeValue(sfr.SitUp.Trim(), "仰臥起坐", student, _sc.Boy_仰臥起坐);
                            sfr.SitAndReachDegree = GetMeValue(sfr.SitAndReach.Trim(), "坐姿體前彎", student, _sc.Boy_坐姿體前彎);
                            sfr.StandingLongJumpDegree = GetMeValue(sfr.StandingLongJump.Trim(), "立定跳遠", student, _sc.Boy_立定跳遠);
                            sfr.CardiorespiratoryDegree = GetMeValue(dotorsec(sfr.Cardiorespiratory.Trim()), "心肺適能", student, _sc.Boy_心肺適能);
                        }
                        else
                        {
                            //如果有內容,則不予處理
                            if (string.IsNullOrEmpty(sfr.SitUpDegree))
                                sfr.SitUpDegree = GetMeValue(sfr.SitUp.Trim(), "仰臥起坐", student, _sc.Boy_仰臥起坐);

                            if (string.IsNullOrEmpty(sfr.SitAndReachDegree))
                                sfr.SitAndReachDegree = GetMeValue(sfr.SitAndReach.Trim(), "坐姿體前彎", student, _sc.Boy_坐姿體前彎);

                            if (string.IsNullOrEmpty(sfr.StandingLongJumpDegree))
                                sfr.StandingLongJumpDegree = GetMeValue(sfr.StandingLongJump.Trim(), "立定跳遠", student, _sc.Boy_立定跳遠);

                            if (string.IsNullOrEmpty(sfr.CardiorespiratoryDegree))
                                sfr.CardiorespiratoryDegree = GetMeValue(dotorsec(sfr.Cardiorespiratory.Trim()), "心肺適能", student, _sc.Boy_心肺適能);
                        }

                        UpdateList.Add(sfr);
                    }
                    else if (student.Gender == "女")
                    {
                        if (CheckPreferred) //2019/6/14 - 擇優覆蓋
                        {
                            //資料是否存在
                            string SitUpDegreeA = sfr.SitUpDegree.Trim();
                            string SitUpDegreeB = GetMeValue(sfr.SitUp, "仰臥起坐", student, _sc.Girl_仰臥起坐);
                            sfr.SitUpDegree = CompareFitness(SitUpDegreeA, SitUpDegreeB);

                            string SitAndReachA = sfr.SitAndReachDegree.Trim();
                            string SitAndReachB = GetMeValue(sfr.SitAndReach, "坐姿體前彎", student, _sc.Girl_坐姿體前彎);
                            sfr.SitAndReachDegree = CompareFitness(SitAndReachA, SitAndReachB);

                            string StandingLongJumpA = sfr.StandingLongJumpDegree.Trim();
                            string StandingLongJumpB = GetMeValue(sfr.StandingLongJump, "立定跳遠", student, _sc.Girl_立定跳遠);
                            sfr.StandingLongJumpDegree = CompareFitness(StandingLongJumpA, StandingLongJumpB);

                            string CardiorespiratoryA = sfr.CardiorespiratoryDegree.Trim();
                            string CardiorespiratoryB = GetMeValue(dotorsec(sfr.Cardiorespiratory), "心肺適能", student, _sc.Girl_心肺適能);
                            sfr.CardiorespiratoryDegree = CompareFitness(CardiorespiratoryA, CardiorespiratoryB);
                        }
                        else if (CheckCover)
                        {
                            sfr.SitUpDegree = GetMeValue(sfr.SitUp, "仰臥起坐", student, _sc.Girl_仰臥起坐);
                            sfr.SitAndReachDegree = GetMeValue(sfr.SitAndReach, "坐姿體前彎", student, _sc.Girl_坐姿體前彎);
                            sfr.StandingLongJumpDegree = GetMeValue(sfr.StandingLongJump, "立定跳遠", student, _sc.Girl_立定跳遠);
                            sfr.CardiorespiratoryDegree = GetMeValue(dotorsec(sfr.Cardiorespiratory), "心肺適能", student, _sc.Girl_心肺適能);
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(sfr.SitUpDegree))
                                sfr.SitUpDegree = GetMeValue(sfr.SitUp, "仰臥起坐", student, _sc.Girl_仰臥起坐);

                            if (string.IsNullOrEmpty(sfr.SitAndReachDegree))
                                sfr.SitAndReachDegree = GetMeValue(sfr.SitAndReach, "坐姿體前彎", student, _sc.Girl_坐姿體前彎);

                            if (string.IsNullOrEmpty(sfr.StandingLongJumpDegree))
                                sfr.StandingLongJumpDegree = GetMeValue(sfr.StandingLongJump, "立定跳遠", student, _sc.Girl_立定跳遠);

                            if (string.IsNullOrEmpty(sfr.CardiorespiratoryDegree))
                                sfr.CardiorespiratoryDegree = GetMeValue(dotorsec(sfr.Cardiorespiratory), "心肺適能", student, _sc.Girl_心肺適能);
                        }


                        UpdateList.Add(sfr);
                    }
                    else
                    {
                        //性別未定將不予處理
                        FitInfo info = new FitInfo(student);
                        info._info = "學生沒有設定性別!!";
                        FitInfoList.Add(info);
                    }
                }
            }

            #region Log處理

            StringBuilder sbLog = new StringBuilder();
            if (CheckCover)
                sbLog.AppendLine("體適能常模轉換:(使用者設定為「常模有值覆蓋」)");
            else
                sbLog.AppendLine("體適能常模轉換:(使用者設定為「常模有值略過」)");

            foreach (string each in _student.Keys)
            {
                KeyBoStudent student = _student[each];

                StudentFitnessRecord sfr1 = student.sfr;
                StudentFitnessRecord sfr_log = student.sfr_Log;

                if (sfr1 != null && sfr_log != null)
                {
                    StringBuilder sb_123 = new StringBuilder();

                    StringBuilder sb_name = new StringBuilder();
                    string AgeString = "";
                    string Age鑑測 = "";

                    if (student.Age != null)
                    {
                        AgeString = string.Format("{0}歲{1}個月又{2}天", student.Age.年, student.Age.月, student.Age.日);
                        Age鑑測 = "" + student.Age.鑑測年齡;
                    }
                    else
                    {
                        AgeString = "無相關資訊";
                        Age鑑測 = "無相關資訊";
                    }

                    sb_name.AppendLine(string.Format("班級「{0}」座號「{1}」學生「{2}」\n實際年齡「{3}」鑑測年齡「{4}」", student.ClassName, student.SeatNo, student.Name, AgeString, Age鑑測));

                    if (sfr1.SitUpDegree != sfr_log.SitUpDegree)
                        sb_123.AppendLine(string.Format("仰臥起坐「{0}」由「{1}」換算為「{2}」", sfr1.SitUp, sfr_log.SitUpDegree, sfr1.SitUpDegree));

                    if (sfr1.SitAndReachDegree != sfr_log.SitAndReachDegree)
                        sb_123.AppendLine(string.Format("坐姿體前彎「{0}」由「{1}」換算為「{2}」", sfr1.SitAndReach, sfr_log.SitAndReachDegree, sfr1.SitAndReachDegree));

                    if (sfr1.StandingLongJumpDegree != sfr_log.StandingLongJumpDegree)
                        sb_123.AppendLine(string.Format("立定跳遠「{0}」由「{1}」換算為「{2}」", sfr1.StandingLongJump, sfr_log.StandingLongJumpDegree, sfr1.StandingLongJumpDegree));

                    if (sfr1.CardiorespiratoryDegree != sfr_log.CardiorespiratoryDegree)
                        sb_123.AppendLine(string.Format("心肺適能「{0}」由「{1}」換算為「{2}」", sfr1.Cardiorespiratory, sfr_log.CardiorespiratoryDegree, sfr1.CardiorespiratoryDegree));

                    sb_123.AppendLine("");

                    sbLog.AppendLine(sb_name.ToString() + sb_123.ToString());

                }
                else
                {

                    FitInfo info = new FitInfo(student);
                    info._info = "沒有體適能資料!";
                    FitInfoList.Add(info);
                }
            }

            #endregion

            if (UpdateList.Count > 0)
            {
                tool._A.UpdateValues(UpdateList);
                FISCA.LogAgent.ApplicationLog.Log("體適能", "常模換算", sbLog.ToString());
            }

            e.Result = FitInfoList;
        }

        private string CompareFitness(string ValueA, string ValueB)
        {
            if (FitnessDic.ContainsKey(ValueA) && FitnessDic.ContainsKey(ValueB))
            {
                //如果新資料大於目前資料
                if (FitnessDic[ValueA] > FitnessDic[ValueB])
                {
                    return ValueA;
                }
                else
                {
                    return ValueB;
                }
            }
            else if (FitnessDic.ContainsKey(ValueA) && !FitnessDic.ContainsKey(ValueB))
            {
                return ValueA;
            }
            else if (!FitnessDic.ContainsKey(ValueA) && FitnessDic.ContainsKey(ValueB))
            {
                //舊資料不存在,新資料比較大
                return ValueB;
            }

            return "";
        }

        private TestAge getAge2(DateTime 生日, DateTime 測驗日期)
        {

            TestAge ta = new TestAge();

            double a = (測驗日期 - 生日).TotalDays;
            ta.年 = int.Parse(Math.Floor(a / 365).ToString());
            ta.鑑測年齡 = int.Parse(Math.Floor(a / 365).ToString());

            double C = a % 365;
            ta.月 = int.Parse(Math.Floor(C / 30).ToString());

            ta.日 = int.Parse(Math.Floor(C % 30).ToString());

            return ta;
        }

        /// <summary>
        /// 心肺適能換算
        /// 傳入字串,並且依據是否有小數點
        /// 來決定是否進行秒數換算
        /// </summary>
        private string dotorsec(string _value)
        {
            if (_value.Contains("."))
            {
                string[] dot = _value.Split('.');
                int x, y, z;
                int.TryParse(dot[0], out x); //分
                int.TryParse(dot[1], out y); //秒
                z = x * 60 + y;
                return z.ToString();
            }
            else
            {
                return _value;
            }
        }


        private string GetMeValue(string fitValue, string fitname, KeyBoStudent student, Dictionary<int, Dictionary<int, string>> dic)
        {
            if (fitValue == "免測")
                return "免測";


            if (student.Age != null)
            {
                if (dic.ContainsKey(student.Age.鑑測年齡))
                {
                    int fitParseIntValue;
                    if (int.TryParse(fitValue, out fitParseIntValue))
                    {
                        if (dic[student.Age.鑑測年齡].ContainsKey(fitParseIntValue))
                        {
                            return dic[student.Age.鑑測年齡][fitParseIntValue];
                        }
                        else
                        {
                            FitInfo info = new FitInfo(student);
                            info._info = string.Format("{0}:「{1}」不在表定範圍!!", fitname, fitValue);
                            FitInfoList.Add(info);
                        }
                    }
                    else
                    {
                        FitInfo info = new FitInfo(student);
                        info._info = string.Format("{0}:「{1}」並非數字!!", fitname, fitValue);
                        FitInfoList.Add(info);
                    }
                }
                else
                {
                    FitInfo info = new FitInfo(student);
                    info._info = string.Format("{0}:鑑測年齡未落在「13~23歲」範圍內", fitname);
                    FitInfoList.Add(info);
                }
            }
            else
            {
                FitInfo info = new FitInfo(student);
                info._info = string.Format("{0}:沒有年齡資料!!", fitname);
                FitInfoList.Add(info);
            }

            return "";
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStart.Enabled = true;

            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    List<FitInfo> FitInfoList = (List<FitInfo>)e.Result;

                    eh(null, EventArgs.Empty);

                    if (FitInfoList.Count > 0)
                    {
                        DialogResult dr = MsgBox.Show("常模轉換作業已完成!\n您是否要檢視詳細訊息!", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);

                        if (dr == System.Windows.Forms.DialogResult.Yes)
                        {
                            InfoForm _if = new InfoForm(FitInfoList);
                            _if.ShowDialog();
                        }
                    }
                    else
                    {
                        MsgBox.Show("常模轉換作業已完成!");
                    }
                }
                else
                {
                    MsgBox.Show("常模轉換作業發生錯誤!!\n" + e.Error.Message);
                }
            }
            else
            {
                MsgBox.Show("常模轉換作業已取消!!");
            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        //{
        //     DialogResult dr = MsgBox.Show("您確定要開啟體適能官方網站?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
        //     if (dr == System.Windows.Forms.DialogResult.Yes)
        //     {
        //          System.Diagnostics.Process.Start("http://www.fitness.org.tw/");
        //     }
        //}

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ViewParse vp = new ViewParse();
            vp.ShowDialog();
        }
    }
}
