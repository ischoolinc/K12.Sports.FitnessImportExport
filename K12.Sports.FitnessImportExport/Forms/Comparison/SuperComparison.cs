using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace K12.Sports.FitnessImportExport
{
     /// <summary>
     /// 常模對照表
     /// </summary>
     class SuperComparison
     {
          /// <summary>
          /// 年齡 : 值 : 常模
          /// </summary>
          public Dictionary<int, Dictionary<int, string>> Boy_仰臥起坐 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Girl_仰臥起坐 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Boy_坐姿體前彎 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Girl_坐姿體前彎 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Boy_立定跳遠 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Girl_立定跳遠 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Boy_心肺適能 { get; set; }
          public Dictionary<int, Dictionary<int, string>> Girl_心肺適能 { get; set; }
          public SuperComparison(XmlElement xml)
          {
               XmlElement Xml_Boy1 = (XmlElement)xml.SelectSingleNode("boy/仰臥起坐");
               Boy_仰臥起坐 = GetParse(Xml_Boy1, true);

               XmlElement Xml_Girl1 = (XmlElement)xml.SelectSingleNode("girl/仰臥起坐");
               Girl_仰臥起坐 = GetParse(Xml_Girl1, true);

               XmlElement Xml_Boy2 = (XmlElement)xml.SelectSingleNode("boy/坐姿體前彎");
               Boy_坐姿體前彎 = GetParse(Xml_Boy2, true);

               XmlElement Xml_Girl2 = (XmlElement)xml.SelectSingleNode("girl/坐姿體前彎");
               Girl_坐姿體前彎 = GetParse(Xml_Girl2, true);

               XmlElement Xml_Boy3 = (XmlElement)xml.SelectSingleNode("boy/立定跳遠");
               Boy_立定跳遠 = GetParse(Xml_Boy3, true);

               XmlElement Xml_Girl3 = (XmlElement)xml.SelectSingleNode("girl/立定跳遠");
               Girl_立定跳遠 = GetParse(Xml_Girl3, true);

               XmlElement Xml_Boy4 = (XmlElement)xml.SelectSingleNode("boy/心肺適能");
               Boy_心肺適能 = GetParse(Xml_Boy4, false);

               XmlElement Xml_Girl4 = (XmlElement)xml.SelectSingleNode("girl/心肺適能");
               Girl_心肺適能 = GetParse(Xml_Girl4, false);
          }

          private Dictionary<int, Dictionary<int, string>> GetParse(XmlElement Xml_Boy, bool Start)
          {
               Dictionary<int, Dictionary<int, string>> dic = new Dictionary<int, Dictionary<int, string>>();
               int StartIndex;
               if (!Start)
                    StartIndex = 9;
               else
                    StartIndex = 7;
               for (int year = StartIndex; year <= 23; year++)
               {
                    if (!dic.ContainsKey(year))
                    {
                         dic.Add(year, new Dictionary<int, string>());
                    }

                    XmlElement each_n = (XmlElement)Xml_Boy.SelectSingleNode("_" + year);
                    if (each_n != null)
                    {
                         string kj = each_n.GetAttribute("請加強");
                         SetValue(year, kj, "請加強", dic, Start);

                         kj = each_n.GetAttribute("中等");
                         SetValue(year, kj, "中等", dic, Start);

                         kj = each_n.GetAttribute("銅牌");
                         SetValue(year, kj, "銅牌", dic, Start);

                         kj = each_n.GetAttribute("銀牌");
                         SetValue(year, kj, "銀牌", dic, Start);

                         kj = each_n.GetAttribute("金牌");
                         SetValue(year, kj, "金牌", dic, Start);
                    }
               }

               return dic;
          }

          private void SetValue(int year, string kj, string name, Dictionary<int, Dictionary<int, string>> dic, bool Start)
          {
               int y, z;
               string[] kjkj = kj.Split(',');
               if (int.TryParse(kjkj[0], out y) && int.TryParse(kjkj[1], out z))
               {
                    if (Start)
                    {
                         for (int x = y; x <= z; x++)
                         {
                              if (!dic[year].ContainsKey(x))
                              {
                                   dic[year].Add(x, name);
                              }
                         }
                    }
                    else
                    {
                         for (int x = y; x >= z; x--)
                         {
                              if (!dic[year].ContainsKey(x))
                              {
                                   dic[year].Add(x, name);
                              }
                         }
                    }
               }
          }
     }
}
