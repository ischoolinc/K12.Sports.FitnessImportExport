
using Aspose.Cells;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Sports.FitnessImportExport.DAO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;


// 2016/7/18 穎驊繼續製作全校體適能統計報表，預計這兩天完成

namespace K12.Sports.FitnessImportExport
{
    public partial class SchoolFitnessStatisticsReport : BaseForm
    {
        public SchoolFitnessStatisticsReport()
        {
            InitializeComponent();

            //動態新增學年選擇，從今年(民國)開始算，往過去減5年
            int ThisSchoolYear = Int32.Parse(K12.Data.School.DefaultSchoolYear);
            for (int i = 0; i < 5; i++)
            {
                comboBox_ChooseSchoolYear.Items.Add(ThisSchoolYear - i);
            };           
        }

        private void buttonX1_Click(object s, EventArgs ea)
        {

            string schoolYear = comboBox_ChooseSchoolYear.Text;

            List<String> Error_List = new List<string>();

            BackgroundWorker BGW = new BackgroundWorker();
            BGW.WorkerReportsProgress = true;

            #region 資料處理

            BGW.DoWork += delegate(object sender, DoWorkEventArgs e)
            {

                BGW.ReportProgress(5);

                // 為列印Excel 先New 物件，注意下行方法只能參考新的 Aspose.Cell_201402，如果用舊的話會有錯誤
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(new MemoryStream(Properties.Resources.全校體適能中等以上_含中等_各項目百分比統計表));

                Cells cs0 = wb.Worksheets[0].Cells;
                
                FISCA.UDT.AccessHelper accessHelper = new FISCA.UDT.AccessHelper();

                Aspose.Words.Document Template;
                Template = new Aspose.Words.Document(new MemoryStream(Properties.Resources.班級體適能確認單範本));
                
              

                // 存放學生ID 與班級名稱的對照
                Dictionary<string, string> studentID_to_className = new Dictionary<string, string>();

                // 存放學生ID 與學生名稱的對照
                Dictionary<string, string> studentID_to_studentName = new Dictionary<string, string>();

                //目前體適能一共五個等級，分別為金牌、銀牌、銅牌、中等、請加強，下面字典為蒐集每班各體適能項目"請加強"的人數使用
                Dictionary<string, to_be_Improve_counter> dic_class_fitness_to_be_Improve = new Dictionary<string, to_be_Improve_counter>();

                // 計算BGW進度使用
                int progress = 0;

                //統計"全校" 體適能項目 "請加強"人數使用
                to_be_Improve_counter Total_school_to_be_Improve_counter = new to_be_Improve_counter();

                //蒐集全校正常在學有班級學生之全部ID
                var studentIDList = new List<string>();

                //蒐集全校正常在學有班級且有"體適能"資料學生之全部ID ，注意此項與studentIDList不一定永遠一樣
                var studentIDList_fitness = new List<string>();

                var studentRecordList = new List<StudentRecord>();

                #region 取得全校班級，並將全校學生做班級分類

                // 取得選取班級，(不給使用者指定選取了，直接選全校)
                //List<ClassRecord> ClassList = K12.Data.Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
                

                studentRecordList = K12.Data.Student.SelectAll();

                foreach (var stuRec in studentRecordList) {
                                        
                    // 0 = 一般生 ，如此一來可以避免選到畢業班，另外也必須要有班級才行
                    if (stuRec.Status == 0 &&stuRec.Class !=null) { 

                    studentIDList.Add(stuRec.ID);

                    studentID_to_className.Add(stuRec.ID, stuRec.Class.Name);

                    studentID_to_studentName.Add(stuRec.ID, stuRec.Name);

                    if (!dic_class_fitness_to_be_Improve.ContainsKey(stuRec.Class.Name)) {

                        dic_class_fitness_to_be_Improve.Add(stuRec.Class.Name, new to_be_Improve_counter());
                                        
                    }
                 

                    }                                               
                }


                // 穎驊筆記，下面的方法註解掉，原因是如果直接選取全校班級Class.SelectAll() ，會選到已經畢業的班級，就現階段來說比較麻煩處理，
                //不如像上面直接找尋全部學生再配給他們班級

                //List<ClassRecord> ClassList = K12.Data.Class.SelectAll();
                //foreach (ClassRecord classrecord in ClassList)
                //{
                //    foreach (var studentRec in classrecord.Students)
                //    {
                //        studentIDList.Add(studentRec.ID);

                //        studentID_to_className.Add(studentRec.ID, classrecord.Name);
                //    }

                //    dic_class_fitness_to_be_Improve.Add(classrecord.Name, new to_be_Improve_counter());


                //} 


                #endregion

                BGW.ReportProgress(20);

                #region 取得體適能資料並分類

                //取得全學生的體適能資料
                var studentFitnessRecordList = accessHelper.Select<StudentFitnessRecord>(string.Format("ref_student_id in ('{0}') AND school_year = {1}", string.Join("','", studentIDList), schoolYear));

                // 把各學生的體適能資料和班級做集合整理
                foreach (var fitnessRec in studentFitnessRecordList)
                {
                    //蒐集全校正常在學有班級且有"體適能"資料學生之全部ID ，注意此項與studentIDList不一定永遠一樣，因為很有可能有在studentIDList的學生卻沒有體適能資料，
                    //自然在取得全學生體適能資料studentFitnessRecordList就不會出現，但我們又必須要將沒有資料的人視為"缺考"、"零分"、"待加強"，所以要再出一個studentIDList_fitness
                    //之後與studentIDList內的ID做比較，找出其班級，把總人數、待加強人數加上去

                    studentIDList_fitness.Add(fitnessRec.StudentID);

                    // 計算坐姿體前彎各班待加強人數

                    if (fitnessRec.SitAndReachDegree == "請加強")
                    {

                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitAndReachDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitAndReachDegree_total++;

                    }
                    // 用來處理假如該學生沒有常模的狀況，此時不能將之算入不及格或是總母群內，另存錯誤提醒視窗
                    
                    //2016/7/21 修正，因為實際拿各個學校資料測試，發現其實真實學校資料都缺蠻多的，會造成錯誤提醒視窗一大包
                    //恩正說學生沒有體適能資料、沒有體適能常模，不是我們的責任，是各個學校應該要自己負責，所以資料不齊者視為"0分"、"缺考"、"不及格"
                    else if (fitnessRec.SitAndReachDegree == "")
                    {
                        Error_List.Add("班級:" + studentID_to_className[fitnessRec.StudentID] + "，" + "學生:" + studentID_to_studentName[fitnessRec.StudentID]+"沒有坐姿體前彎常模資料，將不會納入計算，請確認是否忘記常模計算");

                        //缺體適能常模資料，就當你不合格要"請加強"
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitAndReachDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitAndReachDegree_total++;

                    }
                    else
                    {
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitAndReachDegree_total++;

                    }

                    // 計算立定跳遠各班待加強人數

                    if (fitnessRec.StandingLongJumpDegree == "請加強")
                    {

                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].StandingLongJumpDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].StandingLongJumpDegree_total++;

                    }
                    // 用來處理假如該學生沒有常模的狀況，此時不能將之算入不及格或是總母群內，另存錯誤提醒視窗

                    //2016/7/21 修正，因為實際拿各個學校資料測試，發現其實真實學校資料都缺蠻多的，會造成錯誤提醒視窗一大包
                    //恩正說學生沒有體適能資料、沒有體適能常模，不是我們的責任，是各個學校應該要自己負責，所以資料不齊者視為"0分"、"缺考"、"不及格"
                    else if (fitnessRec.StandingLongJumpDegree == "")
                    {

                        Error_List.Add("班級:" + studentID_to_className[fitnessRec.StudentID] + "，" + "學生:" + studentID_to_studentName[fitnessRec.StudentID] + "沒有立定跳遠常模資料，將不會納入計算，請確認是否忘記常模計算");

                        //缺體適能常模資料，就當你不合格要"請加強"
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].StandingLongJumpDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].StandingLongJumpDegree_total++;

                    }
                    else
                    {
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].StandingLongJumpDegree_total++;

                    }


                    // 計算坐仰臥起坐各班待加強人數

                    if (fitnessRec.SitUpDegree == "請加強")
                    {

                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitUpDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitUpDegree_total++;

                    }
                    // 用來處理假如該學生沒有常模的狀況，此時不能將之算入不及格或是總母群內，另存錯誤提醒視窗

                    //2016/7/21 修正，因為實際拿各個學校資料測試，發現其實真實學校資料都缺蠻多的，會造成錯誤提醒視窗一大包
                    //恩正說學生沒有體適能資料、沒有體適能常模，不是我們的責任，是各個學校應該要自己負責，所以資料不齊者視為"0分"、"缺考"、"不及格"
                    else if (fitnessRec.SitUpDegree == "")
                    {

                        Error_List.Add("班級:" + studentID_to_className[fitnessRec.StudentID] + "，" + "學生:" + studentID_to_studentName[fitnessRec.StudentID] + "沒有仰臥起坐常模資料，將不會納入計算，請確認是否忘記常模計算");

                        //缺體適能常模資料，就當你不合格要"請加強"
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitUpDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitUpDegree_total++;

                    }
                    else
                    {
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].SitUpDegree_total++;

                    }

                    // 計算心肺適能各班待加強人數

                    if (fitnessRec.CardiorespiratoryDegree == "請加強")
                    {

                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].CardiorespiratoryDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].CardiorespiratoryDegree_total++;

                    }
                    // 用來處理假如該學生沒有常模的狀況，此時不能將之算入不及格或是總母群內，另存錯誤提醒視窗

                    //2016/7/21 修正，因為實際拿各個學校資料測試，發現其實真實學校資料都缺蠻多的，會造成錯誤提醒視窗一大包
                    //恩正說學生沒有體適能資料、沒有體適能常模，不是我們的責任，是各個學校應該要自己負責，所以資料不齊者視為"0分"、"缺考"、"不及格"
                    else if (fitnessRec.CardiorespiratoryDegree == "")
                    {

                        Error_List.Add("班級:" + studentID_to_className[fitnessRec.StudentID] + "，" + "學生:" + studentID_to_studentName[fitnessRec.StudentID] + "沒有心肺適能常模資料，將不會納入計算，請確認是否忘記常模計算");

                        //缺體適能常模資料，就當你不合格要"請加強"
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].CardiorespiratoryDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].CardiorespiratoryDegree_total++;

                    }
                    else
                    {
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].CardiorespiratoryDegree_total++;

                    }

                    //2016/11/11 光棍節，穎驊新增計算班級四個項目(坐姿體前彎、立定跳遠、仰臥起坐、心肺適能)都通過比例(在金牌、銀牌、銅牌、中等、待加強五個評等中至少拿中等)

                    // 四大項目都必須要有常模資料，才會進行計算，否則即使只缺一項資料其他項目都通過，也會將之不算四項目都通過
                    if (fitnessRec.SitAndReachDegree != "" && fitnessRec.StandingLongJumpDegree != "" && fitnessRec.SitUpDegree != "" && fitnessRec.CardiorespiratoryDegree != "")
                    {
                        if (fitnessRec.SitAndReachDegree != "請加強" && fitnessRec.StandingLongJumpDegree != "請加強" && fitnessRec.SitUpDegree != "請加強" && fitnessRec.CardiorespiratoryDegree != "請加強")
                        {

                            dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].Four_Item_All_Pass_counter++;
                            dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].Four_Item_All_Pass_counter_total++;

                        }
                        else
                        {
                            dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].Four_Item_All_Pass_counter_total++;

                        }

                    }
                    else 
                    {
                        dic_class_fitness_to_be_Improve[studentID_to_className[fitnessRec.StudentID]].Four_Item_All_Pass_counter_total++;
                    
                    }

                    // 2016/7/22 上一版已在昨天(7/21)早上出去了，今天再做別的東西時發現ReportProgress有點問題，必須要先轉型有浮點數的類型(decimal 、float)算完後，再轉回int
                    //否則 1 / studentFitnessRecordList.Count 一除下來的型別int 可能永遠都是零。由於此功能在實際功能其實沒甚麼影響，先把Code改好，下次等有需求再更新。

                    progress += (int)(((decimal)1 / studentFitnessRecordList.Count) * 60);

                    BGW.ReportProgress(20 + progress);

                }

                //蒐集全校正常在學有班級且有"體適能"資料學生之全部ID ，注意此項與studentIDList不一定永遠一樣，因為很有可能有在studentIDList的學生卻沒有體適能資料，
                //自然在取得全學生體適能資料studentFitnessRecordList就不會出現，但我們又必須要將沒有資料的人視為"缺考"、"零分"、"待加強"，所以要再出一個studentIDList_fitness
                //之後與studentIDList內的ID做比較，找出其班級，把總人數、待加強人數加上去

                // 全校身份一般、有班級的學生
                foreach(var stuID in studentIDList)
                {
                    //卻沒有體適能資料的話，就在該班把每一個項目總人數、待加強人數給加上去
                    if(!studentIDList_fitness.Contains(stuID))

                    {

                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].SitAndReachDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].SitAndReachDegree_total++;

                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].StandingLongJumpDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].StandingLongJumpDegree_total++;

                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].SitUpDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].SitUpDegree_total++;


                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].CardiorespiratoryDegree_failed_counter++;
                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].CardiorespiratoryDegree_total++;


                        //沒有體適能資料，視為沒通過，在四項全過統計直接加分母總人數就好。
                        dic_class_fitness_to_be_Improve[studentID_to_className[stuID]].Four_Item_All_Pass_counter_total++;

                        Error_List.Add("班級:" + studentID_to_className[stuID] + "，" + "學生:" + studentID_to_studentName[stuID] + "沒有體適能資料 請確認是否忘記輸入");
                                                           
                    }
                                                                                                                           
                }
 


                #endregion

                BGW.ReportProgress(80);

                #region 報表填值

                cs0[0, 0].Value = schoolYear + "學年度中等以上(含中等)各項目百分比統計表";

                int RowCounter = 3;

                int ColCounter = 0;

                int endRow = dic_class_fitness_to_be_Improve.Count + 3;

                //看看全資料有沒有錯誤，如果有錯，全校的百分比將不計算，而顯示錯誤
                //2016/7/21 因應恩正所說，我們其實不太需要為殘破的資料負責任，所以這個Bool 會暫時用不到
                bool DataBroken = false;

                Workbook template = new Workbook();

                template = new Aspose.Cells.Workbook(new MemoryStream(Properties.Resources.全校體適能中等以上_含中等_各項目百分比統計表_程式用樣版_));

                // 固定複製另一份的樣板，最後一行"全校"的那一行，這樣不論班級有多少，最後一行都會是全校
                cs0.CopyRow(template.Worksheets[0].Cells, 12, endRow);


                foreach (var item in dic_class_fitness_to_be_Improve)
                {

                    cs0.CopyRow(template.Worksheets[0].Cells, 3, RowCounter);

                    // 第一欄填班級
                    cs0[RowCounter, ColCounter].Value = item.Key;

                    // 第二欄坐姿體前彎，且分母總數不可為0
                    if (item.Value.SitAndReachDegree_total != 0)
                    {
                        cs0[RowCounter, ColCounter + 1].Value = Math.Round((100 - (item.Value.SitAndReachDegree_failed_counter / item.Value.SitAndReachDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";


                        Total_school_to_be_Improve_counter.SitAndReachDegree_failed_counter += item.Value.SitAndReachDegree_failed_counter;

                        Total_school_to_be_Improve_counter.SitAndReachDegree_total += item.Value.SitAndReachDegree_total;
                    }
                    else
                    {

                        Error_List.Add("班級:" + item.Key + "全班無同學有坐姿體前彎體適能紀錄常模，將造成計算百分比錯誤，請檢查是否忘記常模計算，或是該班同學都沒有體適能資料");

                        // 下行為舊的處理方式。
                        //cs0[RowCounter, ColCounter + 1].Value = "無資料";

                        //新的處理方式，全班數不到人直接0%
                        cs0[RowCounter, ColCounter + 1].Value = "0%";

                        DataBroken = true;

                    }

                    // 第三欄填立定跳遠，且分母總數不可為0
                    if (item.Value.StandingLongJumpDegree_total != 0)
                    {
                        cs0[RowCounter, ColCounter + 2].Value = Math.Round((100 - (item.Value.StandingLongJumpDegree_failed_counter / item.Value.StandingLongJumpDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                        Total_school_to_be_Improve_counter.StandingLongJumpDegree_failed_counter += item.Value.StandingLongJumpDegree_failed_counter;

                        Total_school_to_be_Improve_counter.StandingLongJumpDegree_total += item.Value.StandingLongJumpDegree_total;
                    }
                    else
                    {

                        Error_List.Add("班級:" + item.Key + "全班無同學有立定跳遠體適能紀錄常模，將造成計算百分比錯誤，請檢查是否忘記常模計算，或是該班同學都沒有體適能資料");


                        // 下行為舊的處理方式。
                        //cs0[RowCounter, ColCounter + 2].Value = "無資料";

                        //新的處理方式，全班數不到人直接0%
                        cs0[RowCounter, ColCounter + 2].Value = "0%";

                        DataBroken = true;

                    }


                    // 第四欄填仰臥起坐，且分母總數不可為0
                    if (item.Value.SitUpDegree_total != 0)
                    {
                        cs0[RowCounter, ColCounter + 3].Value = Math.Round((100 - (item.Value.SitUpDegree_failed_counter / item.Value.SitUpDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";


                        Total_school_to_be_Improve_counter.SitUpDegree_failed_counter += item.Value.SitUpDegree_failed_counter;

                        Total_school_to_be_Improve_counter.SitUpDegree_total += item.Value.SitUpDegree_total;

                    }
                    else
                    {
                        Error_List.Add("班級:" + item.Key + "全班無同學有仰臥起坐體適能紀錄常模，將造成計算百分比錯誤，請檢查是否忘記常模計算，或是該班同學都沒有體適能資料");


                        // 下行為舊的處理方式。
                        //cs0[RowCounter, ColCounter + 3].Value = "無資料";

                        //新的處理方式，全班數不到人直接0%
                        cs0[RowCounter, ColCounter + 3].Value = "0%";

                        DataBroken = true;

                    }


                    // 第五欄填心肺適能，且分母總數不可為0
                    if (item.Value.CardiorespiratoryDegree_total != 0)
                    {
                        cs0[RowCounter, ColCounter + 4].Value = Math.Round((100 - (item.Value.CardiorespiratoryDegree_failed_counter / item.Value.CardiorespiratoryDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";


                        Total_school_to_be_Improve_counter.CardiorespiratoryDegree_failed_counter += item.Value.CardiorespiratoryDegree_failed_counter;

                        Total_school_to_be_Improve_counter.CardiorespiratoryDegree_total += item.Value.CardiorespiratoryDegree_total;
                    }

                    else
                    {

                        Error_List.Add("班級:" + item.Key + "全班無同學有心肺體適能紀錄常模，將造成計算百分比錯誤，請檢查是否忘記常模計算，或是該班同學都沒有體適能資料");


                        // 下行為舊的處理方式。
                        //cs0[RowCounter, ColCounter + 4].Value = "無資料";



                        //新的處理方式，全班數不到人直接0%
                        cs0[RowCounter, ColCounter + 4].Value = "0%";



                        DataBroken = true;

                    }

                    //2016/11/11 光棍節，穎驊新增   第六欄填四項皆通過人數百分比統計，且分母總數不可為0
                    if (item.Value.Four_Item_All_Pass_counter_total != 0)
                    {
                        cs0[RowCounter, ColCounter + 5].Value = Math.Round(( item.Value.Four_Item_All_Pass_counter/ item.Value.Four_Item_All_Pass_counter_total) * 100, 0, MidpointRounding.AwayFromZero) + "%";


                        Total_school_to_be_Improve_counter.Four_Item_All_Pass_counter += item.Value.Four_Item_All_Pass_counter;

                        Total_school_to_be_Improve_counter.Four_Item_All_Pass_counter_total += item.Value.Four_Item_All_Pass_counter_total;
                    }


                    RowCounter++;

                }

                cs0[endRow, ColCounter + 1].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.SitAndReachDegree_failed_counter / Total_school_to_be_Improve_counter.SitAndReachDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                cs0[endRow, ColCounter + 2].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.StandingLongJumpDegree_failed_counter / Total_school_to_be_Improve_counter.StandingLongJumpDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                cs0[endRow, ColCounter + 3].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.SitUpDegree_failed_counter / Total_school_to_be_Improve_counter.SitUpDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                cs0[endRow, ColCounter + 4].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.CardiorespiratoryDegree_failed_counter / Total_school_to_be_Improve_counter.CardiorespiratoryDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";


                //新增全校四項通過百分比
                cs0[endRow, ColCounter + 5].Value = Math.Round((Total_school_to_be_Improve_counter.Four_Item_All_Pass_counter / Total_school_to_be_Improve_counter.Four_Item_All_Pass_counter_total* 100), 0, MidpointRounding.AwayFromZero) + "%";



                // 下行為舊的處理方式。
                //if (!DataBroken == true)
                //{
                //    cs0[endRow, ColCounter + 1].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.SitAndReachDegree_failed_counter / Total_school_to_be_Improve_counter.SitAndReachDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                //    cs0[endRow, ColCounter + 2].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.StandingLongJumpDegree_failed_counter / Total_school_to_be_Improve_counter.StandingLongJumpDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                //    cs0[endRow, ColCounter + 3].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.SitUpDegree_failed_counter / Total_school_to_be_Improve_counter.SitUpDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                //    cs0[endRow, ColCounter + 4].Value = Math.Round((100 - (Total_school_to_be_Improve_counter.CardiorespiratoryDegree_failed_counter / Total_school_to_be_Improve_counter.CardiorespiratoryDegree_total) * 100), 0, MidpointRounding.AwayFromZero) + "%";

                //}
                //else
                //{

                //    cs0[endRow, ColCounter + 1].Value = "無法計算";

                //    cs0[endRow, ColCounter + 2].Value = "無法計算";

                //    cs0[endRow, ColCounter + 3].Value = "無法計算";

                //    cs0[endRow, ColCounter + 4].Value = "無法計算";                                                                                                                                                                                                                           

                //}
                
                #endregion
                e.Result = wb;
                BGW.ReportProgress(100);
                
            }; 
            #endregion


            #region 計算DoWork完成百分比

            BGW.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級體適能通知單產生中...", e.ProgressPercentage);
            };
            
            #endregion

            #region 列印存檔

            BGW.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {


                //2016/7/21 不再顯示錯誤訊息

                // 顯示錯誤訊息
                //if (Error_List.Count > 0)
                //{

                //    StringBuilder sb = new StringBuilder();

                //    foreach (var errorMsg in Error_List)
                //    {
                //        sb.AppendLine(errorMsg);

                //    }

                //    MsgBox.Show(sb.ToString());

                //}

                #region RunWorkerCompleted


                Workbook workbook = e.Result as Workbook;

                if (workbook == null)
                    return;


                // 以後記得存Excel 都用新版的Xlsx，可以避免ㄧ些不必要的問題(EX: sheet 只能到1023張)
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "另存新檔";
                save.FileName = "全校體適能中等以上(含中等)各項目百分比統計表"+"("+schoolYear+"學年度"+")";
                save.Filter = "Excel檔案 (*.Xlsx)|*.Xlsx|所有檔案 (*.*)|*.*";

                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        workbook.Save(save.FileName, Aspose.Cells.SaveFormat.Xlsx);
                        System.Diagnostics.Process.Start(save.FileName);
                    }
                    catch
                    {
                        MessageBox.Show("檔案儲存失敗");


                    }
                }
                #endregion
            };
            
            #endregion

            BGW.RunWorkerAsync();
            this.Close();

        }
        // 離開
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
