using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Sports.FitnessImportExport.UDT;
using FISCA.UDT;

namespace K12.Sports.FitnessImportExport.UIForm
{
    public partial class SetCardiorespiratoryTest : BaseForm
    {
        List<FitnessConfigRecord> ConfigList;
        FitnessConfigRecord SelectRecord;
        string configItemName = "心肺耐力施測方式";

        // 2024/10/17，經過討論將心肺耐力顯示改成 800/1600公尺跑走
        string SelItem1Value = "心肺耐力";
        string SelItem1Display = "800/1600公尺跑走";

        public SetCardiorespiratoryTest()
        {
            InitializeComponent();
            ConfigList = new List<FitnessConfigRecord>();
            SelectRecord = new FitnessConfigRecord();
        }

        private void SetCardiorespiratoryTest_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;

            // 學年度
            int defSchoolYear;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out defSchoolYear))
            {
                for (int i = (defSchoolYear + 2); i > (defSchoolYear - 3); i--)
                    cboSchoolYear.Items.Add(i);

                cboSchoolYear.Text = defSchoolYear + "";
            }

            // 項目名稱
            cboSelItem.Items.Add(SelItem1Display);
            cboSelItem.Items.Add("漸速耐力跑");
            cboSelItem.Items.Add("");

            cboSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;
            cboSelItem.DropDownStyle = ComboBoxStyle.DropDownList;

            // 取得設定資料
            GetConfigData();

            GetItemValueBySchoolYear(defSchoolYear);
        }

        private void GetItemValueBySchoolYear(int SchoolYear)
        {
            cboSelItem.Text = "";
            SelectRecord = null;
            foreach (FitnessConfigRecord fc in ConfigList)
            {
                if (fc.SchoolYear == SchoolYear)
                {
                    if (fc.ItemValue == SelItem1Value)
                        cboSelItem.Text = SelItem1Display;
                    else
                        cboSelItem.Text = fc.ItemValue;

                    SelectRecord = fc;
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void GetConfigData()
        {
            AccessHelper accessHelper = new AccessHelper();
            List<FitnessConfigRecord> dataList = accessHelper.Select<FitnessConfigRecord>();
            ConfigList.Clear();
            foreach (FitnessConfigRecord data in dataList)
            {
                if (data.ItemName == configItemName)
                    ConfigList.Add(data);
            }

        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            try
            {
                // 檢查學年度是否輸入
                int sy;
                bool isSyError = false;
                if (!int.TryParse(cboSchoolYear.Text, out sy))
                {
                    isSyError = true;
                }

                if (isSyError)
                {
                    MsgBox.Show("請輸入學年度!");
                    btnSave.Enabled = true;
                    return;
                }

                if (SelectRecord == null)
                {
                    SelectRecord = new FitnessConfigRecord();
                    SelectRecord.ItemName = configItemName;
                }

                SelectRecord.SchoolYear = sy;
                if (cboSelItem.Text == SelItem1Display)
                    SelectRecord.ItemValue = SelItem1Value;
                else
                    SelectRecord.ItemValue = cboSelItem.Text;
                SelectRecord.Save();
                MsgBox.Show("儲存完成");
                this.Close();
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }


            btnSave.Enabled = true;
        }

        private void cboSelItem_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            int sy;
            cboSelItem.Text = "";
            SelectRecord = null;
            if (int.TryParse(cboSchoolYear.Text, out sy))
            {
                GetItemValueBySchoolYear(sy);
            }
        }
    }
}
