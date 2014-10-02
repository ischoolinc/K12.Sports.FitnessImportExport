using Aspose.Words;
using Aspose.Words.Drawing;
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

namespace K12.Sports.FitnessImportExport.Report
{
    public partial class FitnessProveSingle : BaseForm
    {
        /// <summary>
        /// 體適能證明單設定檔
        /// </summary>
        private string CadreConfig = "K12.Sports.FitnessImportExport.FitnessProveSingle.cs";

        private BackgroundWorker BGW = new BackgroundWorker();

        int 記錄多少筆 = 8;

        public FitnessProveSingle()
        {
            InitializeComponent();
        }

        private void FitnessProveSingle_Load(object sender, EventArgs e)
        {
            BGW.DoWork += BGW_DoWork;
            BGW.RunWorkerCompleted += BGW_RunWorkerCompleted;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (BGW.IsBusy)
            {
                MsgBox.Show("忙碌中,稍後再試!!");
                return;
            }

            btnPrint.Enabled = false;
            BGW.RunWorkerAsync();

        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {

            #region 範本

            //整理取得報表範本
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(CadreConfig);
            Aspose.Words.Document Template;

            if (ConfigurationInCadre.Template == null)
            {
                //如果範本為空,則建立一個預設範本
                Campus.Report.ReportConfiguration ConfigurationInCadre_1 = new Campus.Report.ReportConfiguration(CadreConfig);
                ConfigurationInCadre_1.Template = new Campus.Report.ReportTemplate(Properties.Resources.體適能證明單_範本, Campus.Report.TemplateType.Word);
                Template = ConfigurationInCadre_1.Template.ToDocument();
            }
            else
            {
                //如果已有範本,則取得樣板
                Template = ConfigurationInCadre.Template.ToDocument();
            }

            #endregion

            List<StudentRecord> StudentList = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);

            //取得資料
            List<string> StudentIDList = new List<string>();
            foreach (StudentRecord stud in StudentList)
            {
                StudentIDList.Add(stud.ID);
            }

            Dictionary<string, StudentRSRecord> AllSRSRDic = GetRSR(StudentList, StudentIDList);

            //填資料部份
            DataTable table = new DataTable();

            table.Columns.Add("學校名稱");
            table.Columns.Add("列印日期");
            table.Columns.Add("校長名稱");

            table.Columns.Add("班級");
            table.Columns.Add("座號");
            table.Columns.Add("學號");
            table.Columns.Add("姓名");

            table.Columns.Add("新生照片1");
            table.Columns.Add("新生照片2");
            table.Columns.Add("畢業照片1");
            table.Columns.Add("畢業照片2");
            SerColumn(table, "學年度");
            SerColumn(table, "測驗日期");

            SerColumn(table, "身高");
            SerColumn(table, "身高常模");

            SerColumn(table, "體重");
            SerColumn(table, "體重常模");

            SerColumn(table, "坐姿體前彎");
            SerColumn(table, "坐姿體前彎常模");

            SerColumn(table, "立定跳遠");
            SerColumn(table, "立定跳遠常模");

            SerColumn(table, "仰臥起坐");
            SerColumn(table, "仰臥起坐常模");

            SerColumn(table, "心肺適能");
            SerColumn(table, "心肺適能常模");

            foreach (string studentID in AllSRSRDic.Keys)
            {
                StudentRSRecord student = AllSRSRDic[studentID];

                DataRow row = table.NewRow();
                row["學校名稱"] = K12.Data.School.ChineseName;

                row["列印日期"] = DateTime.Today.ToShortDateString();

                XmlElement xml = K12.Data.School.Configuration["學校資訊"].PreviousData;

                if (xml.SelectSingleNode("ChancellorEnglishName") != null)
                {
                    row["校長名稱"] = xml.SelectSingleNode("ChancellorChineseName").InnerText;
                }
                else
                {
                    row["校長名稱"] = "";
                }

                row["班級"] = student._student.Class != null ? student._student.Class.Name : "";
                row["座號"] = student._student.SeatNo.HasValue ? student._student.SeatNo.Value.ToString() : "";
                row["學號"] = student._student.StudentNumber;
                row["姓名"] = student._student.Name;

                row["新生照片1"] = student.學生入學照片;
                row["新生照片2"] = student.學生入學照片;
                row["畢業照片1"] = student.學生畢業照片;
                row["畢業照片2"] = student.學生畢業照片;

                student._ResultList.Sort(SortResultScore);

                int y = 1;
                foreach (StudentFitnessRecord Fitness in student._ResultList)
                {
                    if (y <= 記錄多少筆)
                    {
                        row[string.Format("學年度{0}", y)] = Fitness.SchoolYear.ToString();
                        row[string.Format("測驗日期{0}", y)] = Fitness.TestDate.ToShortDateString();

                        row[string.Format("身高{0}", y)] = Fitness.Height;
                        row[string.Format("身高常模{0}", y)] = Fitness.HeightDegree;

                        row[string.Format("體重{0}", y)] = Fitness.Weight;
                        row[string.Format("體重常模{0}", y)] = Fitness.WeightDegree;

                        row[string.Format("坐姿體前彎{0}", y)] = Fitness.Weight;
                        row[string.Format("坐姿體前彎常模{0}", y)] = Fitness.WeightDegree;

                        row[string.Format("立定跳遠{0}", y)] = Fitness.Weight;
                        row[string.Format("立定跳遠常模{0}", y)] = Fitness.WeightDegree;

                        row[string.Format("仰臥起坐{0}", y)] = Fitness.Weight;
                        row[string.Format("仰臥起坐常模{0}", y)] = Fitness.WeightDegree;

                        row[string.Format("心肺適能{0}", y)] = Fitness.Weight;
                        row[string.Format("心肺適能常模{0}", y)] = Fitness.WeightDegree;
                        y++;
                    }
                }

                table.Rows.Add(row);
            }

            Document PageOne = (Document)Template.Clone(true);
            PageOne.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
            PageOne.MailMerge.Execute(table);
            e.Result = PageOne;
        }

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            if (e.FieldName == "新生照片1" || e.FieldName == "新生照片2")
            {
                #region 新生照片
                if (!string.IsNullOrEmpty(e.FieldValue.ToString()))
                {
                    byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                    if (photo != null && photo.Length > 0)
                    {
                        DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                        photoBuilder.MoveToField(e.Field, true);
                        e.Field.Remove();
                        //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                        Shape photoShape = new Shape(e.Document, ShapeType.Image);
                        photoShape.ImageData.SetImage(photo);
                        photoShape.WrapType = WrapType.Inline;
                        //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                        //cell.CellFormat.LeftPadding = 0;
                        //cell.CellFormat.RightPadding = 0;
                        if (e.FieldName == "新生照片1")
                        {
                            // 1吋
                            photoShape.Width = ConvertUtil.MillimeterToPoint(25);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(35);
                        }
                        else
                        {
                            //2吋
                            photoShape.Width = ConvertUtil.MillimeterToPoint(35);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(45);
                        }
                        //paragraph.AppendChild(photoShape);
                        photoBuilder.InsertNode(photoShape);
                    }
                }
                #endregion
            }
            else if (e.FieldName == "畢業照片1" || e.FieldName == "畢業照片2")
            {
                #region 畢業照片
                if (!string.IsNullOrEmpty(e.FieldValue.ToString()))
                {
                    byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                    if (photo != null && photo.Length > 0)
                    {
                        DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                        photoBuilder.MoveToField(e.Field, true);
                        e.Field.Remove();
                        //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                        Shape photoShape = new Shape(e.Document, ShapeType.Image);
                        photoShape.ImageData.SetImage(photo);
                        photoShape.WrapType = WrapType.Inline;
                        //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                        //cell.CellFormat.LeftPadding = 0;
                        //cell.CellFormat.RightPadding = 0;
                        if (e.FieldName == "畢業照片1")
                        {
                            // 1吋
                            photoShape.Width = ConvertUtil.MillimeterToPoint(25);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(35);
                        }
                        else
                        {
                            //2吋
                            photoShape.Width = ConvertUtil.MillimeterToPoint(35);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(45);
                        }
                        //paragraph.AppendChild(photoShape);
                        photoBuilder.InsertNode(photoShape);
                    }
                }
                #endregion
            }
        }

        public int SortResultScore(StudentFitnessRecord rsr1, StudentFitnessRecord rsr2)
        {
            return rsr1.SchoolYear.CompareTo(rsr2.SchoolYear);
        }

        private void SerColumn(DataTable table, string n)
        {
            for (int x = 1; x <= 記錄多少筆; x++)
            {
                table.Columns.Add(string.Format("{0}{1}", n, x));
            }
        }

        /// <summary>
        /// 學生資料整理
        /// </summary>
        private Dictionary<string, StudentRSRecord> GetRSR(List<StudentRecord> StudentList, List<string> StudentIDList)
        {
            Dictionary<string, StudentRSRecord> dic = new Dictionary<string, StudentRSRecord>();

            List<StudentFitnessRecord> RSList = tool._A.Select<StudentFitnessRecord>(string.Format("ref_student_id in ('{0}')", string.Join("','", StudentIDList)));

            //整理學生基本資料記錄
            foreach (StudentRecord stud in StudentList)
            {
                StudentRSRecord rsr = new StudentRSRecord(stud);
                if (!dic.ContainsKey(stud.ID))
                {
                    dic.Add(stud.ID, rsr);
                }
            }

            //整理學生社團記錄
            foreach (StudentFitnessRecord rsr in RSList)
            {
                if (dic.ContainsKey(rsr.StudentID))
                {
                    dic[rsr.StudentID].SetRSR(rsr);
                }
            }

            // 入學照片
            Dictionary<string, string> _PhotoPDict = new Dictionary<string, string>();

            // 畢業照片
            Dictionary<string, string> _PhotoGDict = new Dictionary<string, string>();

            if (StudentList.Count != 0)
            {
                // 入學照片
                _PhotoPDict = K12.Data.Photo.SelectFreshmanPhoto(StudentIDList);

                // 畢業照片
                _PhotoGDict = K12.Data.Photo.SelectGraduatePhoto(StudentIDList);
            }

            //處理照片
            foreach (string studnetID in dic.Keys)
            {
                if (_PhotoPDict.ContainsKey(studnetID))
                {
                    dic[studnetID].學生入學照片 = _PhotoPDict[studnetID];
                }

                if (_PhotoGDict.ContainsKey(studnetID))
                {
                    dic[studnetID].學生畢業照片 = _PhotoGDict[studnetID];
                }
            }

            return dic;
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            btnPrint.Enabled = true;

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

                        SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        SaveFileDialog1.FileName = "體適能個人證明單";

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

                    this.Close();
                }
                else
                {
                    MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                }
            }


        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lbTempAll_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "體適能證明單_合併欄位總表.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.體適能個人證明單_功能變數總表, 0, Properties.Resources.體適能個人證明單_功能變數總表.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(CadreConfig);
            //畫面內容(範本內容,預設樣式
            Campus.Report.TemplateSettingForm TemplateForm;
            if (ConfigurationInCadre.Template != null)
            {
                TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(Properties.Resources.體適能證明單_範本, Campus.Report.TemplateType.Word));
            }
            else
            {
                ConfigurationInCadre.Template = new Campus.Report.ReportTemplate(Properties.Resources.體適能證明單_範本, Campus.Report.TemplateType.Word);
                TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(Properties.Resources.體適能證明單_範本, Campus.Report.TemplateType.Word));
            }

            //預設名稱
            TemplateForm.DefaultFileName = "體適能證明單(範本)";
            //如果回傳為OK
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                //設定後樣試,回傳
                ConfigurationInCadre.Template = TemplateForm.Template;
                //儲存
                ConfigurationInCadre.Save();
            }
        }
    }
}
