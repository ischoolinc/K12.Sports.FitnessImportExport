using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using K12.Presentation;
using System.ComponentModel;
using FISCA.Permission;
using Campus.DocumentValidator;

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

            // 加入"匯出"按鈕以及圖示
            NLDPanels.Student.RibbonBarItems["體適能"]["匯出"].Image = Properties.Resources.Export_Image;
            NLDPanels.Student.RibbonBarItems["體適能"]["匯出"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;

            // 加入"匯出體適能"按鈕
            FISCA.Presentation.MenuButton btnExport = NLDPanels.Student.RibbonBarItems["體適能"]["匯出"]["匯出體適能"];
            // 設定權限
            btnExport.Enable = Permissions.IsEnableFitnessExport;
            // 設定動作
            btnExport.Click += delegate
            {
                //SmartSchool.API.PlugIn.Export.Exporter exporter = new Actions.ExportStudentFitness();
                //Actions.FrmFitnessExport wizard = new Actions.FrmFitnessExport(exporter.Text, exporter.Image);
                //exporter.InitializeExport(wizard);
                //wizard.ShowDialog();
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

            // 加入"匯出"按鈕以及圖示
            NLDPanels.Student.RibbonBarItems["體適能"]["匯入"].Image = Properties.Resources.Import_Image;
            NLDPanels.Student.RibbonBarItems["體適能"]["匯入"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;

            // 加入"匯入體適能"按鈕
            FISCA.Presentation.MenuButton btnImport = NLDPanels.Student.RibbonBarItems["體適能"]["匯入"]["匯入體適能"];
            // 設定權限
            btnImport.Enable = Permissions.IsEnableFitnessImport;
            // 設定動作
            btnImport.Click += delegate
            {
                // 準備所有一般生的學生ID, 之後驗證資料時會用到
                Global._AllStudentNumberIDTemp = DAO.FDQuery.GetAllStudenNumberDict();

                //Actions.FrmFitnessImportBaseForm frm = new Actions.FrmFitnessImportBaseForm();
                //frm.ShowDialog();
                ImportExport.ImportStudentFitness frmImport = new ImportExport.ImportStudentFitness();
                frmImport.Execute();
            };


            // 在權限畫面出現"體適能資料項目"權限
            Catalog catalog1 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog1.Add(new DetailItemFeature(Permissions.KeyFitnessContent, "體適能"));

            // 在權限畫面出現"匯出體適能"權限
            Catalog catalog2 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog2.Add(new RibbonFeature(Permissions.KeyFitnessExport, "匯出體適能"));

            // 在權限畫面出現"匯入體適能"權限
            Catalog catalog3 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog3.Add(new RibbonFeature(Permissions.KeyFitnessImport, "匯入體適能"));
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
