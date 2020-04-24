using Aspose.Cells;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace K12.Sports.FitnessImportExport
{
    public partial class InfoForm : BaseForm
    {

        List<FitInfo> _list { get; set; }

        public InfoForm(List<FitInfo> list)
        {
            InitializeComponent();
            _list = list;
            _list.Sort(SortClass);
            dataGridViewX1.DataSource = _list;
        }

        private int SortClass(FitInfo x, FitInfo y)
        {
            string xx = x._class.PadLeft(10, '0');
            string yy = y._class.PadLeft(10, '0');

            xx += x._seatno.PadLeft(3, '0');
            yy += y._seatno.PadLeft(3, '0');

            xx += x._name.PadLeft(10, '0');
            yy += y._name.PadLeft(10, '0');

            return xx.CompareTo(yy);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            List<string> j_line = new List<string>();
            foreach (FitInfo each in _list)
            {
                if (!j_line.Contains(each._studentID))
                {
                    j_line.Add(each._studentID);
                }
            }
            K12.Presentation.NLDPanels.Student.RemoveFromTemp(K12.Presentation.NLDPanels.Student.TempSource);
            K12.Presentation.NLDPanels.Student.AddToTemp(j_line);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "常模轉換訊息";
            saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
            DataTable dt = new DataTable();

            dt.Columns.Add("班級");
            dt.Columns.Add("座號");
            dt.Columns.Add("姓名");
            dt.Columns.Add("訊息");

            foreach (FitInfo each in _list)
            {
                DataRow row = dt.NewRow();
                row["班級"] = each._class;
                row["座號"] = each._seatno;
                row["姓名"] = each._name;
                row["訊息"] = each._info;
                dt.Rows.Add(row);
            }
            wb.Worksheets[0].Cells.ImportDataTable(dt, true, "A1");
            wb.Save(saveFileDialog1.FileName, FileFormatType.Xlsx);
            System.Diagnostics.Process.Start(saveFileDialog1.FileName);
        }

        private void DataGridViewX1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }

    public class FitInfo
    {
        public FitInfo(KeyBoStudent s)
        {
            _studentID = s.ID;
            _class = s.ClassName;
            _seatno = s.SeatNo;
            _name = s.Name;
        }

        public string _studentID { get; set; }
        public string _class { get; set; }

        public string _seatno { get; set; }

        public string _name { get; set; }

        public string _info { get; set; }

    }
}
