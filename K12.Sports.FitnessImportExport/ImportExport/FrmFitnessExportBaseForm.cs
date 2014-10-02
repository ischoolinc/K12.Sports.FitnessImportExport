using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using K12.Data;
using Aspose.Cells;
using System.IO;

namespace K12.Sports.FitnessImportExport.ImportExport
{
    public partial class FrmFitnessExportBaseForm : BaseForm
    {
        private readonly string _Title = "匯出體適能";

        private readonly int _MAX_ROW_COUNT = 65535;
        private readonly int _START_ROW = 16;

        public FrmFitnessExportBaseForm()
        {
            InitializeComponent();
        }

        private void FrmFitnessExportBaseForm_Load(object sender, EventArgs e)
        {
            integerInput1.Text = K12.Data.School.DefaultSchoolYear;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            this.btnExport.Enabled = false;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "另存新檔";
            saveFileDialog1.FileName = "" + _Title + "(" + integerInput1.Value + "學年度).xls";
            saveFileDialog1.Filter = "Excel (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // 新增背景執行緒來處理資料的匯出
                BackgroundWorker BGW = new BackgroundWorker();
                BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
                BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

                // 把檔案儲存的路徑當作參數傳入
                BGW.RunWorkerAsync(new object[] { saveFileDialog1.FileName });
            }
            else
            {
                this.btnExport.Enabled = true;
            }
        }

        // 主要邏輯區塊
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)((object[])e.Argument)[0];

            #region 取得需要的資料
            int SchoolYear = integerInput1.Value;
            bool ExportDegree = ckExportDegree.Checked;

            // 取得選取的學生ID
            List<string> studentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;

            // 取得學生的基本資料包括班級資料
            List<DAO.StudentInfo> studentRecords = DAO.FDQuery.GetStudnetInfoByIDList(studentIDList);

            // 取得學生的體適能資料
            List<DAO.StudentFitnessRecord> FitnessRecords = DAO.StudentFitness.SelectByStudentIDListAndSchoolYear(studentIDList, SchoolYear);

            Dictionary<string, DAO.StudentFitnessRecord> FitnessRecordsDic = new Dictionary<string, DAO.StudentFitnessRecord>();
            foreach (DAO.StudentFitnessRecord each in FitnessRecords)
            {
                if (!FitnessRecordsDic.ContainsKey(each.StudentID))
                {
                    FitnessRecordsDic.Add(each.StudentID, each);
                }
            }

            // Excel的表頭
            string[] ExcelColumnNames;
            if (ExportDegree == true)
                ExcelColumnNames = Global._ExcelDataDegreeTitle;
            else
                ExcelColumnNames = Global._ExcelDataTitle;
            #endregion

            // 所有資料得集合
            List<ExcelRowRecord> excelRowRecords = new List<ExcelRowRecord>();

            #region 把資料組合起來
            // 學生
            foreach (DAO.StudentInfo student in studentRecords)
            {
                if (FitnessRecordsDic.ContainsKey(student.Student_ID))
                {
                    DAO.StudentFitnessRecord fitnessRecord = FitnessRecordsDic[student.Student_ID];

                    // 設定輸出的欄位
                    ExcelRowRecord rec = new ExcelRowRecord(ExcelColumnNames);
                    // 設定每個欄位的值
                    rec.SetDataForExport(student, fitnessRecord, ExportDegree);
                    excelRowRecords.Add(rec);
                }
                else
                {
                    // 設定輸出的欄位
                    ExcelRowRecord rec = new ExcelRowRecord(ExcelColumnNames);
                    // 設定每個欄位的值
                    rec.SetDataForExport(student, null, ExportDegree);
                    excelRowRecords.Add(rec);
                }

            }   // end of foreach (StudentRecord student in studentRecords)

            #endregion

            // 排序
            excelRowRecords.Sort(SortData);

            // 開起現有的樣板檔案
            Workbook report = new Workbook();
            MemoryStream ms = new MemoryStream(Properties.Resources.體適能資料上傳格式);
            report.Open(ms);

            Worksheet sheet = report.Worksheets[0];
            sheet.Name = Global._SheetName;

            // 輸出表頭
            int RowIndex = _START_ROW - 1;
            int colIndex = 0;
            foreach (string columnName in ExcelColumnNames)
            {
                sheet.Cells[RowIndex, colIndex++].PutValue(columnName);
            }

            //填入資料
            RowIndex = _START_ROW;
            foreach (ExcelRowRecord excelRowRecord in excelRowRecords)
            {
                if (RowIndex > _MAX_ROW_COUNT)
                {
                    break;
                }

                SetDataDetail(sheet, excelRowRecord, RowIndex);
                RowIndex++;
            }

            // 儲存結果
            e.Result = new object[] { report, fileName, RowIndex > _MAX_ROW_COUNT };

        }

        // 當背景程式結束, 就會呼叫method
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.btnExport.Enabled = true;

            if (e.Error == null)
            {
                Workbook report = (Workbook)((object[])e.Result)[0];
                bool overLimit = (bool)((object[])e.Result)[2];

                #region 儲存 Excel
                string path = (string)((object[])e.Result)[1];

                if (File.Exists(path))
                {
                    bool needCount = true;
                    try
                    {
                        File.Delete(path);
                        needCount = false;
                    }
                    catch { }
                    int i = 1;
                    while (needCount)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                        else
                        {
                            try
                            {
                                File.Delete(newPath);
                                path = newPath;
                                break;
                            }
                            catch { }
                        }
                    }
                }
                try
                {
                    File.Create(path).Close();
                }
                catch
                {
                    SaveFileDialog sd = new SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = Path.GetFileNameWithoutExtension(path) + ".xls";
                    sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            File.Create(sd.FileName);
                            path = sd.FileName;
                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                report.Save(path, FileFormatType.Excel2003);
                #endregion
                if (overLimit)
                    MsgBox.Show("匯出資料已經超過Excel的極限(65536筆)。\n超出的資料無法被匯出。\n\n請減少選取學生人數。");
                System.Diagnostics.Process.Start(path);
            }
            else
            {
                // SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage(_Title + "發生未預期錯誤。");
                MsgBox.Show(_Title + "發生未預期錯誤。\n" + e.Error.Message);
            }
        }

        #region Excel處理
        // 對Excel新增資料
        private void SetDataDetail(Worksheet sheet, ExcelRowRecord rec, int RowIndex)
        {
            int columnIndex = 0;
            foreach (string columnName in rec._Columns)
            {
                sheet.Cells[RowIndex, columnIndex++].PutValue(rec.GetColumnValue(columnName));
            }
        }

        #endregion

        #region 排序的方法

        /// <summary>
        /// 排序:年級/班級序號/班級名稱/學號/姓名
        /// </summary>
        private int SortData(ExcelRowRecord obj1, ExcelRowRecord obj2)
        {
            string seatno1 = obj1.GetColumnValue("年級").PadLeft(1, '0');       // 年級
            seatno1 += obj1.GetOthersValue("班級序號").PadLeft(3, '0');         // 班級序號
            seatno1 += obj1.GetColumnValue("班級名稱").PadLeft(20, '0');        // 班級名稱
            seatno1 += obj1.GetColumnValue("學號/座號").PadLeft(15, '0');       // 學號
            seatno1 += obj1.GetOthersValue("姓名").PadLeft(10, '0');            // 姓名

            string seatno2 = obj2.GetColumnValue("年級").PadLeft(1, '0');       // 年級
            seatno2 += obj2.GetOthersValue("班級序號").PadLeft(3, '0');         // 班級序號
            seatno2 += obj2.GetColumnValue("班級名稱").PadLeft(20, '0');        // 班級名稱
            seatno2 += obj2.GetColumnValue("學號/座號").PadLeft(15, '0');       // 學號
            seatno2 += obj2.GetOthersValue("姓名").PadLeft(10, '0');            // 姓名

            return seatno1.CompareTo(seatno2);
        }

        #endregion

        #region 自訂類別
        internal class ExcelRowRecord
        {
            public string[] _Columns;

            // key: column name; value: column value
            public Dictionary<string, string> _ColumnValue;

            public Dictionary<string, string> _OthersValue;

            /// <summary>
            /// 欄位名稱請不要重複!!
            /// </summary>
            /// <param name="columns"></param>
            public ExcelRowRecord(string[] columns)
            {
                _Columns = columns;
                _ColumnValue = new Dictionary<string, string>();

                foreach (string columnName in _Columns)
                {
                    _ColumnValue.Add(columnName, "");
                }

                _OthersValue = new Dictionary<string, string>();
            }

            /// <summary>
            /// 依照欄位名稱設定欄位內容
            /// </summary>
            /// <param name="columnName"></param>
            /// <param name="columnValue"></param>
            public void SetColumnValue(string columnName, string columnValue)
            {
                if (_ColumnValue.ContainsKey(columnName))
                {
                    _ColumnValue[columnName] = columnValue;
                }
            }

            /// <summary>
            /// 依照欄位名稱取得欄位內容
            /// </summary>
            /// <param name="columnName"></param>
            /// <returns></returns>
            public string GetColumnValue(string columnName)
            {
                if (_ColumnValue.ContainsKey(columnName))
                {
                    return _ColumnValue[columnName];
                }

                return "";
            }

            /// <summary>
            /// 依照欄位名稱設定欄位內容
            /// </summary>
            /// <param name="columnName"></param>
            /// <param name="columnValue"></param>
            public void SetOthersValue(string columnName, string columnValue)
            {
                if (_OthersValue.ContainsKey(columnName))
                {
                    _OthersValue[columnName] = columnValue;
                }
                else
                {
                    _OthersValue.Add(columnName, columnValue);
                }
            }

            /// <summary>
            /// 依照欄位名稱取得欄位內容
            /// </summary>
            /// <param name="columnName"></param>
            /// <returns></returns>
            public string GetOthersValue(string columnName)
            {
                if (_OthersValue.ContainsKey(columnName))
                {
                    return _OthersValue[columnName];
                }

                return "";
            }

            /// <summary>
            /// 主要是給匯出時用的
            /// </summary>
            public void SetDataForExport(DAO.StudentInfo studentRecord, DAO.StudentFitnessRecord fitnessRecord, bool isExportDegree)
            {
                if (fitnessRecord != null)
                {
                    // 測驗日期
                    SetColumnValue("測驗日期", Utility.ConvertDateTimeToChineseDateTime(fitnessRecord.TestDate));
                }
                else
                {
                    // 測驗日期
                    SetColumnValue("測驗日期", string.Empty);
                }

                // 學校類別
                SetColumnValue("學校類別", "國中");

                // 年級
                if (studentRecord.Class_Grade_Year == "1")
                {
                    SetColumnValue("年級", "7");
                }
                else if (studentRecord.Class_Grade_Year == "2")
                {
                    SetColumnValue("年級", "8");
                }
                else if (studentRecord.Class_Grade_Year == "3")
                {
                    SetColumnValue("年級", "9");
                }
                else
                {
                    SetColumnValue("年級", studentRecord.Class_Grade_Year);
                }

                // 班級名稱
                SetColumnValue("班級名稱", studentRecord.Class_Class_Name);

                // 班級序號 for sort
                SetOthersValue("班級序號", studentRecord.Class_Display_Order);

                // 學號/座號
                SetColumnValue("學號/座號", studentRecord.Student_Student_Number);

                // 性別
                SetColumnValue("性別", studentRecord.Student_Gender);

                // 身分證字號
                SetColumnValue("身分證字號", studentRecord.Student_ID_Number);

                // 生日
                SetColumnValue("生日", studentRecord.Student_Birthday);

                if (fitnessRecord != null)
                {
                    // 身高
                    SetColumnValue("身高", fitnessRecord.Height);

                    // 體重
                    SetColumnValue("體重", fitnessRecord.Weight);

                    // 坐姿體前彎
                    SetColumnValue("坐姿體前彎", fitnessRecord.SitAndReach);

                    // 立定跳遠
                    SetColumnValue("立定跳遠", fitnessRecord.StandingLongJump);

                    // 仰臥起坐
                    SetColumnValue("仰臥起坐", fitnessRecord.SitUp);

                    // 心肺適能
                    SetColumnValue("心肺適能", fitnessRecord.Cardiorespiratory);


                    // 姓名 for sort
                    SetOthersValue("姓名", studentRecord.Student_Name);

                    if (isExportDegree == true)
                    {
                        // 身高常模
                        SetColumnValue("身高常模", fitnessRecord.HeightDegree);

                        // 體重常模
                        SetColumnValue("體重常模", fitnessRecord.WeightDegree);

                        // 坐姿體前彎常模
                        SetColumnValue("坐姿體前彎常模", fitnessRecord.SitAndReachDegree);

                        // 立定跳遠常模
                        SetColumnValue("立定跳遠常模", fitnessRecord.StandingLongJumpDegree);

                        // 仰臥起坐常模
                        SetColumnValue("仰臥起坐常模", fitnessRecord.SitUpDegree);

                        // 心肺適能常模
                        SetColumnValue("心肺適能常模", fitnessRecord.CardiorespiratoryDegree);
                    }
                }
                else
                {
                    // 身高
                    SetColumnValue("身高", string.Empty);

                    // 體重
                    SetColumnValue("體重", string.Empty);

                    // 坐姿體前彎
                    SetColumnValue("坐姿體前彎", string.Empty);

                    // 立定跳遠
                    SetColumnValue("立定跳遠", string.Empty);

                    // 仰臥起坐
                    SetColumnValue("仰臥起坐", string.Empty);

                    // 心肺適能
                    SetColumnValue("心肺適能", string.Empty);


                    // 姓名 for sort
                    SetOthersValue("姓名", studentRecord.Student_Name);

                    if (isExportDegree == true)
                    {
                        // 身高常模
                        SetColumnValue("身高常模", string.Empty);

                        // 體重常模
                        SetColumnValue("體重常模", string.Empty);

                        // 坐姿體前彎常模
                        SetColumnValue("坐姿體前彎常模", string.Empty);

                        // 立定跳遠常模
                        SetColumnValue("立定跳遠常模", string.Empty);

                        // 仰臥起坐常模
                        SetColumnValue("仰臥起坐常模", string.Empty);

                        // 心肺適能常模
                        SetColumnValue("心肺適能常模", string.Empty);
                    }
                }
            }
        }
        #endregion

    }
}
