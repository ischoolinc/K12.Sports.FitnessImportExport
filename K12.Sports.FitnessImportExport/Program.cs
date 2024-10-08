﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using K12.Presentation;
using System.ComponentModel;
using FISCA.Permission;
using Campus.DocumentValidator;
using K12.Sports.FitnessImportExport.Report;
using FISCA.Presentation.Controls;
using FISCA.Presentation;

namespace K12.Sports.FitnessImportExport
{
    public class Program
    {

        [FISCA.MainMethod]
        public static void Main()
        {
            // 檢查UDT是否存在
            CheckUDTExist();

            #region 自訂驗證規則

            FactoryProvider.RowFactory.Add(new ValidationRule.FitnessRowValidatorFactory());

            #endregion

            // 把"體適能資料"加入資料項目
            if (FISCA.Permission.UserAcl.Current[Permissions.KeyFitnessContent].Editable || FISCA.Permission.UserAcl.Current[Permissions.KeyFitnessContent].Viewable)
                K12.Presentation.NLDPanels.Student.AddDetailBulider<DetailContents.StudentFitnessContent>();
            // 2018.09.22 [ischoolKingdom] Vicky依據 [J學務][01] 體適能功能重覆整理、UI調整 項目，將體適能()的"()"去除，統一命名為體適能。
            RibbonBarItem FitnessBar = NLDPanels.Student.RibbonBarItems["體適能"];


            // 加入"匯出"按鈕以及圖示
            FitnessBar["匯出"].Image = Properties.Resources.Export_Image;
            FitnessBar["匯出"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;

            // 加入"匯出"按鈕以及圖示
            FitnessBar["匯入"].Image = Properties.Resources.Import_Image;
            FitnessBar["匯入"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;

            FitnessBar["報表"].Image = Properties.Resources.paste_64;
            FitnessBar["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;

            FitnessBar["常模轉換"].Image = Properties.Resources.amplify_wave_64;
            FitnessBar["常模轉換"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            FitnessBar["常模轉換"].Enable = false;
            FitnessBar["常模轉換"].Click += delegate
            {
                ComparisonForm cf = new ComparisonForm();
                cf.ShowDialog();
            };

            // 加入"匯出體適能"按鈕
            FISCA.Presentation.MenuButton btnExport2 = FitnessBar["匯出"]["匯出體適能上傳檔"];
            // 設定權限
            btnExport2.Enable = false;
            // 設定動作
            btnExport2.Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    ImportExport.FrmFitnessExportBaseForm frm = new ImportExport.FrmFitnessExportBaseForm();
                    frm.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生!");
                }
            };

            // 加入"匯入體適能"按鈕
            FISCA.Presentation.MenuButton btnImport = FitnessBar["匯入"]["匯入體適能上傳檔"];
            // 設定權限
            btnImport.Enable = Permissions.IsEnableFitnessImport;
            // 設定動作
            btnImport.Click += delegate
            {
                MsgBox.Show("說明:\n1.匯入檔案必須將說明內容(1行~15行)刪除.\n(只需保留標題以下內容)\n2.必須手動增加學年度欄位\n3.匯入功能之鍵值為[身分證字號+學年度]");
                // 準備所有一般生的學生ID, 之後驗證資料時會用到
                Global._AllStudentIDNumberIDTemp = DAO.FDQuery.GetAllIDNumberDict();

                ImportExport.ImportStudentFitness frmImport = new ImportExport.ImportStudentFitness();
                frmImport.Execute();
            };

            FISCA.Presentation.MenuButton btnReport = FitnessBar["報表"]["體適能證明單"];
            btnReport.Enable = false;
            btnReport.Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    FitnessProveSingle fps = new FitnessProveSingle();
                    fps.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生!");
                }
            };

            FISCA.Presentation.MenuButton btnCTest = FitnessBar["設定"]["管理心肺耐力施測方式"];
            btnCTest.Enable = Permissions.IsEnableFitnessTestSetting;
            btnCTest.Click += delegate
            {
                UIForm.SetCardiorespiratoryTest sct = new UIForm.SetCardiorespiratoryTest();
                sct.ShowDialog();
            };



            //2016/6/2 穎驊新增
            //2016/8/3 因應繼斌要求 將"班級體適能通知單" 更名為 "班級體適能確認單"
            FISCA.Presentation.MenuButton ClassFitnessInformReport = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"];
            ClassFitnessInformReport["班級體適能確認單"].Enable = false;
            ClassFitnessInformReport["班級體適能確認單"].Click += delegate
            {
                if (NLDPanels.Class.SelectedSource.Count > 0)
                {
                    ClassFitnessInformReport CFIR = new ClassFitnessInformReport();
                    CFIR.ShowDialog();

                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請先選擇班級!");
                }
            };



            //2016/7/7 穎驊新增，恩正說把"全校體適能統計百分比報表"功能放在學務作業裡面              
            FISCA.Presentation.MenuButton SchoolFitnessStatisticsReport = MotherForm.RibbonBarItems["學務作業", "資料統計"]["報表"]["全校體適能統計百分比報表"];

            SchoolFitnessStatisticsReport.Enable = Permissions.IsEnableSchoolFitnessStatisticsReport;

            SchoolFitnessStatisticsReport.Click += delegate
            {


                SchoolFitnessStatisticsReport SFSR = new SchoolFitnessStatisticsReport();
                SFSR.ShowDialog();



            };


            NLDPanels.Student.SelectedSourceChanged += delegate
            {
                bool check = NLDPanels.Student.SelectedSource.Count > 0;


                //體適能證明單
                btnReport.Enable = check && Permissions.IsEnableFitnessProveSingle;

                //匯出體適能
                btnExport2.Enable = check && Permissions.IsEnableFitnessExport;



                FitnessBar["常模轉換"].Enable = check && Permissions.體適能常模轉換權限;
            };


            //2016/6/2 穎驊新增
            //2016/8/3 因應繼斌要求 將"班級體適能通知單" 更名為 "班級體適能確認單"
            NLDPanels.Class.SelectedSourceChanged += delegate
            {

                bool check = NLDPanels.Class.SelectedSource.Count > 0;

                ClassFitnessInformReport["班級體適能確認單"].Enable = check && Permissions.IsEnableClassFitnessInformReport;

            };


            string 體適能證明單 = "ischool/國中系統/學生/報表/體適能/體適能證明單";
            FISCA.Features.Register(體適能證明單, arg =>
            {
                FitnessProveSingle fps = new FitnessProveSingle();
                fps.ShowDialog();
            });

            // 在權限畫面出現"體適能資料項目"權限
            Catalog catalog1 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog1.Add(new DetailItemFeature(Permissions.KeyFitnessContent, "體適能"));

            // 在權限畫面出現"匯出體適能"權限
            Catalog catalog2 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog2.Add(new RibbonFeature(Permissions.KeyFitnessExport, "匯出體適能"));

            // 在權限畫面出現"匯入體適能"權限
            Catalog catalog3 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog3.Add(new RibbonFeature(Permissions.KeyFitnessImport, "匯入體適能"));

            // 在權限畫面出現"匯出體適能"權限
            Catalog catalog4 = RoleAclSource.Instance["學生"]["報表"];
            catalog4.Add(new RibbonFeature(Permissions.KeyFitnessProveSingle, "體適能證明單"));

            // 在權限畫面出現"常模轉換"權限
            Catalog catalog5 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog5.Add(new RibbonFeature(Permissions.體適能常模轉換, "常模轉換"));


            // 2016//6/2 穎驊新增 ， 在權限畫面出現"班級體適能通知單"權限
            //2016/8/3 因應繼斌要求 將"班級體適能通知單" 更名為 "班級體適能確認單"
            Catalog catalog6 = RoleAclSource.Instance["班級"]["報表"];
            catalog6.Add(new RibbonFeature(Permissions.KeyClassFitnessInformReport, "班級體適能確認單"));


            // 2016//7/7 穎驊新增 ， 在權限畫面出現"全校體適能統計百分比報表"權限(而且是在學務作業)
            Catalog catalog7 = RoleAclSource.Instance["學務作業"]["功能按鈕"];
            catalog7.Add(new RibbonFeature(Permissions.KeySchoolFitnessStatisticsReport, "全校體適能統計百分比報表"));

            // 管理心肺耐力施測方式
            Catalog catalog8 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog8.Add(new RibbonFeature(Permissions.KeyEnableFitnessTestSetting, "管理心肺耐力施測方式"));
        }

        private static void CheckUDTExist()
        {
            // 檢查UDT
            BackgroundWorker bkWork;

            bkWork = new BackgroundWorker();
            bkWork.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            bkWork.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bkWork_RunWorkerCompleted);
            bkWork.RunWorkerAsync();
        }

        static void _bkWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 當有錯誤訊息顯示
            if (Global._ErrorMessageList.Length > 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show(Global._ErrorMessageList.ToString());
            }
        }

        static void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                // 檢查並建立UDT Table
                DAO.StudentFitness.CreateFitnessUDTTable();
            }
            catch (Exception ex)
            {
                Global._ErrorMessageList.AppendLine("載入體適能發生錯誤：" + ex.Message);
            }
        }
    }
}
