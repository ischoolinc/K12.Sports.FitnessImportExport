using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
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
    public partial class ViewParse : BaseForm
    {
        string name1 = "請加強";
        string name2 = "中等";
        string name3 = "銅牌";
        string name4 = "銀牌";
        string name5 = "金牌";

        public ViewParse()
        {
            InitializeComponent();
        }

        private void ViewParse_Load(object sender, EventArgs e)
        {
            try
            {

                //把體適能換算,具體呈現內容
                XmlElement xml = DSXmlHelper.LoadXml(Properties.Resources.Sports_Fitness_Comparison);
                List<DataGridViewRow> rowList = new List<DataGridViewRow>();

                XmlNode Boy_坐姿體前彎 = xml.SelectSingleNode("boy/坐姿體前彎");
                rowList = GetRow("男", "坐姿體前彎", 10, Boy_坐姿體前彎);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                XmlNode Boy_立定跳遠 = xml.SelectSingleNode("boy/立定跳遠");
                rowList = GetRow("男", "立定跳遠", 10, Boy_立定跳遠);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                XmlNode Boy_仰臥起坐 = xml.SelectSingleNode("boy/仰臥起坐");
                rowList = GetRow("男", "仰臥起坐", 10, Boy_仰臥起坐);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                // 2024 改名 心肺耐力
                XmlNode Boy_心肺適能 = xml.SelectSingleNode("boy/心肺適能");
                rowList = GetRow("男", "心肺耐力", 10, Boy_心肺適能);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                // 2024 新增：
                XmlNode Boy_仰臥捲腹 = xml.SelectSingleNode("boy/仰臥捲腹");
                rowList = GetRow("男", "仰臥捲腹", 10, Boy_仰臥捲腹);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                XmlNode Boy_漸速耐力跑 = xml.SelectSingleNode("boy/漸速耐力跑");
                rowList = GetRow("男", "漸速耐力跑", 10, Boy_漸速耐力跑);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                //===================

                XmlNode Girl_坐姿體前彎 = xml.SelectSingleNode("girl/坐姿體前彎");
                rowList = GetRow("女", "坐姿體前彎", 10, Girl_坐姿體前彎);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                XmlNode Girl_立定跳遠 = xml.SelectSingleNode("girl/立定跳遠");
                rowList = GetRow("女", "立定跳遠", 10, Girl_立定跳遠);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                XmlNode Girl_仰臥起坐 = xml.SelectSingleNode("girl/仰臥起坐");
                rowList = GetRow("女", "仰臥起坐", 10, Girl_仰臥起坐);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                // 改名 心肺耐力
                XmlNode Girl_心肺適能 = xml.SelectSingleNode("girl/心肺適能");
                rowList = GetRow("女", "心肺耐力", 10, Girl_心肺適能);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                // 2024 新增：
                XmlNode Girl_仰臥捲腹 = xml.SelectSingleNode("girl/仰臥捲腹");
                rowList = GetRow("女", "仰臥捲腹", 10, Girl_仰臥捲腹);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());

                XmlNode Girl_漸速耐力跑 = xml.SelectSingleNode("girl/漸速耐力跑");
                rowList = GetRow("女", "漸速耐力跑", 10, Girl_漸速耐力跑);
                dataGridViewX1.Rows.AddRange(rowList.ToArray());
            }
            catch (Exception ex)
            {
                MsgBox.Show("解析失敗，" + ex.Message);
            }

        }

        private List<DataGridViewRow> GetRow(string k, string p, int year, XmlNode boy_仰臥起坐)
        {
            List<DataGridViewRow> rowList = new List<DataGridViewRow>();

            for (int x = year; x <= 23; x++)
            {
                foreach (XmlNode year_node in boy_仰臥起坐.SelectNodes("_" + x))
                {
                    if (year_node != null)
                    {
                        string value1 = ((XmlElement)year_node).GetAttribute(name1);
                        string value2 = ((XmlElement)year_node).GetAttribute(name2);
                        string value3 = ((XmlElement)year_node).GetAttribute(name3);
                        string value4 = ((XmlElement)year_node).GetAttribute(name4);
                        string value5 = ((XmlElement)year_node).GetAttribute(name5);

                        DataGridViewRow row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX1);
                        row.Cells[0].Value = k;
                        row.Cells[1].Value = x;
                        row.Cells[2].Value = p;
                        row.Cells[3].Value = value1.Split(',')[0] + "至" + value1.Split(',')[1];

                        // 請加強改成待加強
                        row.Cells[4].Value = "待加強";
                        row.Cells[5].Value = value2.Split(',')[0] + "至" + value2.Split(',')[1];
                        row.Cells[6].Value = name2;
                        row.Cells[7].Value = value3.Split(',')[0] + "至" + value3.Split(',')[1];
                        row.Cells[8].Value = name3;
                        row.Cells[9].Value = value4.Split(',')[0] + "至" + value4.Split(',')[1];
                        row.Cells[10].Value = name4;
                        row.Cells[11].Value = value5.Split(',')[0] + "至" + value5.Split(',')[1];
                        row.Cells[12].Value = name5;


                        rowList.Add(row);
                    }
                }
            }
            return rowList;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
