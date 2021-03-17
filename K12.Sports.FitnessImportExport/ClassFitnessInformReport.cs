
using Aspose.Words;
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

namespace K12.Sports.FitnessImportExport
{

    // 2016/6/1 穎驊製作體適能班級通知單，預計三天內完成
    public partial class ClassFitnessInformReport : BaseForm
    {

        public ClassFitnessInformReport()
        {
            InitializeComponent();

            //動態新增學年選擇，從今年(民國)開始算，往過去減5年
            int ThisSchoolYear = Int32.Parse(K12.Data.School.DefaultSchoolYear);
            for (int i = 0; i < 5; i++)
            {
                comboBox_ChooseSchoolYear.Items.Add(ThisSchoolYear - i);
            };
        }


        // 列印
        private void btnPrint_Click_1(object s, EventArgs ea)
        {

            if (comboBox_ChooseSchoolYear.SelectedItem == null)
            {
                // 請使用者一定要選擇學年度，否則系統會當機
                MsgBox.Show("請先選擇學年度");
            }
            else
            {

                string schoolYear = comboBox_ChooseSchoolYear.Text;
                string returnDate = textBoxHandInDay.Text;
                string printDate = DateTime.Today.ToShortDateString();

                BackgroundWorker BGW = new BackgroundWorker();
                BGW.WorkerReportsProgress = true;

                BGW.DoWork += delegate(object sender, DoWorkEventArgs e)
                {
                    #region DoWork
                    FISCA.UDT.AccessHelper accessHelper = new FISCA.UDT.AccessHelper();

                    Aspose.Words.Document Template;
                    Template = new Aspose.Words.Document(new MemoryStream(Properties.Resources.班級體適能確認單範本1));
                    // 取得選取班級
                    List<ClassRecord> ClassList = K12.Data.Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
                    Dictionary<string, StudentFitnessRecord> dicStudentFitnessRecord = new Dictionary<string, StudentFitnessRecord>();
                    var studentIDList = new List<string>();
                    foreach (ClassRecord classrecord in ClassList)
                    {
                        foreach (var studentRec in classrecord.Students)
                        {
                            studentIDList.Add(studentRec.ID);
                        }
                    }
                    BGW.ReportProgress(10);

                    var studentFitnessRecordList = accessHelper.Select<StudentFitnessRecord>(string.Format("ref_student_id in ('{0}') AND school_year = {1}", string.Join("','", studentIDList), schoolYear));

                    foreach (var fitnessRec in studentFitnessRecordList)
                    {
                        //2021/3/16 -  如果沒有新增,避免爆掉
                        //- By Dylan
                        if (!dicStudentFitnessRecord.ContainsKey(fitnessRec.StudentID))
                        {
                            dicStudentFitnessRecord.Add(fitnessRec.StudentID, fitnessRec);
                        }
                        else
                        {
                            StudentRecord stud = K12.Data.Student.SelectByID(fitnessRec.StudentID);
                            MsgBox.Show(string.Format("學生「{0}」體適能資料重複\n(一學年僅會有一筆體適能紀錄)", stud.Name));
                        }
                    }

                    BGW.ReportProgress(20);

                    //填資料部份
                    DataTable table = new DataTable();
                    table.Columns.Add("製表日期");
                    table.Columns.Add("學年");
                    table.Columns.Add("學期");
                    table.Columns.Add("班級");
                    table.Columns.Add("導師");
                    table.Columns.Add("繳回日期");


                    int classIndex = 0;
                    foreach (ClassRecord classRec in ClassList)
                    {

                        DataRow row = table.NewRow();
                        row["學年"] = schoolYear;

                        row["班級"] = classRec.Name;

                        if (classRec.Teacher != null)
                        {
                            row["導師"] = classRec.Teacher.Name;
                        }
                        //  取得視窗輸入的繳回日期
                        row["繳回日期"] = returnDate;

                        row["製表日期"] = printDate;

                        int studentCounter = 0;

                        foreach (StudentRecord studentRec in classRec.Students)
                        {

                            //2016/11/11 穎驊更正，限制抓取"一般"狀態的學生，要不然會在同一班 抓到畢業、休學、刪除的學生資料
                            if (studentRec.Status == StudentRecord.StudentStatus.一般)
                            {
                                string col = "";

                                col = string.Format("姓名{0}", studentCounter);
                                if (!table.Columns.Contains(col))
                                    table.Columns.Add(col);
                                row[col] = studentRec.Name;

                                col = string.Format("座號{0}", studentCounter);
                                if (!table.Columns.Contains(col))
                                    table.Columns.Add(col);
                                row[col] = studentRec.SeatNo;

                                if (dicStudentFitnessRecord.ContainsKey(studentRec.ID))
                                {
                                    col = string.Format("測驗日期{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].TestDate.ToShortDateString();

                                    col = string.Format("身高{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].Height;

                                    col = string.Format("體重{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].Weight;

                                    col = string.Format("坐姿體前彎{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].SitAndReach;

                                    col = string.Format("坐姿體前彎常模{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].SitAndReachDegree;

                                    col = string.Format("立定跳遠{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].StandingLongJump;

                                    col = string.Format("立定跳遠常模{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].StandingLongJumpDegree;

                                    col = string.Format("仰臥起坐{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].SitUp;

                                    col = string.Format("仰臥起坐常模{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].SitUpDegree;

                                    col = string.Format("心肺適能{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].Cardiorespiratory;

                                    col = string.Format("心肺適能常模{0}", studentCounter);
                                    if (!table.Columns.Contains(col))
                                        table.Columns.Add(col);
                                    row[col] = dicStudentFitnessRecord[studentRec.ID].CardiorespiratoryDegree;
                                }

                                studentCounter++;

                                //2016/11/11 光棍節，穎驊新增，由於目前Word樣板只支援38個學生，當班級學生數量將會有印不下的問題，
                                //因此 將第39位後的學生資料，印在第二頁、第三頁...
                                if (studentCounter >= 38)
                                {
                                    studentCounter = 0;
                                    table.Rows.Add(row);

                                    row = table.NewRow();

                                    row["學年"] = schoolYear;

                                    row["班級"] = classRec.Name;

                                    if (classRec.Teacher != null)
                                    {
                                        row["導師"] = classRec.Teacher.Name;
                                    }
                                    //  取得視窗輸入的繳回日期
                                    row["繳回日期"] = returnDate;

                                    row["製表日期"] = printDate;


                                }
                            }
                        }
                        // 一個row 一班
                        if (studentCounter != 0) 
                        {
                            table.Rows.Add(row);
                        }
                        
                        classIndex++;
                        BGW.ReportProgress(20 + classIndex * 80 / ClassList.Count);

                    }

                    #region 自動生成功變數代碼(開發用很方便，平常註解掉)

                    // 雖然已經講過了，但穎驊不得不大力推薦，這~真~的~超~級~好~用~的!!! 原本自己手動改，三個小時還不一全部改得完、正確，
                    // 用程式自動產生功能變數mailmerge名稱後，十分鐘內就完成&確認檢查完畢了


                    //Document doc = new Document();
                    //DocumentBuilder bu = new DocumentBuilder(doc);
                    //bu.MoveToDocumentStart();
                    //bu.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    //Table table1 = bu.StartTable();

                    //List<string> fitnessItem = new List<string>();

                    //fitnessItem.Add("座號");
                    //fitnessItem.Add("姓名");
                    //fitnessItem.Add("測驗日期");
                    //fitnessItem.Add("身高");
                    //fitnessItem.Add("體重");
                    //fitnessItem.Add("坐姿體前彎");
                    //fitnessItem.Add("坐姿體前彎常模");
                    //fitnessItem.Add("立定跳遠");
                    //fitnessItem.Add("立定跳遠常模");
                    //fitnessItem.Add("仰臥起坐");
                    //fitnessItem.Add("仰臥起坐常模");
                    //fitnessItem.Add("心肺適能");
                    //fitnessItem.Add("心肺適能常模");

                    //    foreach (String item in fitnessItem)
                    //    {
                    //        for (int fitnessCounter = 0; fitnessCounter < 40; fitnessCounter++)
                    //        {
                    //        bu.InsertCell();
                    //        bu.CellFormat.Width = 15;
                    //        bu.InsertField("MERGEFIELD " + item + fitnessCounter + @" \* MERGEFORMAT", "«»");
                    //        bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    //        bu.InsertCell();
                    //        bu.CellFormat.Width = 125;
                    //        bu.Write(item + fitnessCounter);
                    //        bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    //        bu.EndRow();

                    //    }

                    //}
                    //    table1.AllowAutoFit = false;
                    //bu.EndTable();
                    //Document PageOne = (Document)Template.Clone(true);
                    //PageOne = doc;

                    # endregion

                    Document PageOne = (Document)Template.Clone(true);
                    PageOne.MailMerge.Execute(table);
                    PageOne.MailMerge.DeleteFields();


                    e.Result = PageOne;
                    #endregion
                };

                BGW.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("班級體適能通知單產生中...", e.ProgressPercentage);
                };

                BGW.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
                {
                    #region RunWorkerCompleted
                    if (e.Cancelled)
                    {
                        MsgBox.Show("作業已被中止!!");
                    }
                    else
                    {
                        if (e.Error == null)
                        {
                            Document inResult = (Document)e.Result;

                            try
                            {
                                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                                SaveFileDialog1.Filter = "Word (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                                SaveFileDialog1.FileName = "班級體適能通知單";

                                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                                {
                                    inResult.Save(SaveFileDialog1.FileName);
                                    Process.Start(SaveFileDialog1.FileName);
                                }
                                else
                                {
                                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                                    return;
                                }
                            }
                            catch
                            {
                                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                                return;
                            }

                            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級體適能通知單產生完成", 100);
                        }
                        else
                        {
                            MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                        }
                    }
                    #endregion
                };

                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級體適能通知單產生中...", 0);

                BGW.RunWorkerAsync();
                this.Close();
            }
        }
    }

}




