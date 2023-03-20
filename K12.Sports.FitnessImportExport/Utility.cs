using Campus.DocumentValidator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport
{
    class Utility
    {
        /// <summary>
        /// 透過輸入的欄位名稱, 取的匯入資料的值
        /// </summary>
        public static string GetIRowValueString(IRowStream row, string name)
        {
            if(row.Contains(name))
            {
                if(string.IsNullOrEmpty(row.GetValue(name)))
                    return "";

                return row.GetValue(name).Trim();
            }
            else
                return "";
        }

        /// <summary>
        /// 透過輸入的欄位名稱, 取的匯入資料的值, 並轉成DateTime
        /// </summary>
        /// <param name="row"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static DateTime? GetIRowValueDateTime(IRowStream row, string name)
        {
            if (row.Contains(name))
            {
                if (string.IsNullOrEmpty(row.GetValue(name)))
                    return null;

                DateTime dt;
                if (DateTime.TryParse(row.GetValue(name).Trim(), out dt))
                    return dt;
                else
                    return null;
            }
            else
                return null;
        }

        /// <summary>
        /// 透過輸入的欄位名稱, 取的匯入資料的值, 並轉成int
        /// </summary>
        /// <param name="row"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static int? GetIRowValueInt(IRowStream row, string name)
        {
            if (row.Contains(name))
            {
                if (string.IsNullOrEmpty(row.GetValue(name)))
                    return null;
                int retVal;
                if (int.TryParse(row.GetValue(name).Trim(), out retVal))
                    return retVal;
                else
                    return null;
            }
            else
                return null;
        }

        /// <summary>
        ///  把DateTime轉成民國年, 格式: yyyMMdd, 月跟日會補0到兩碼
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ConvertDateTimeToChineseDateTime(DateTime? dt)
        {
            
            if(dt.HasValue)
            {
                DateTime tmp = DateTime.Parse(dt.Value.ToString());

                int year = (tmp.Year) - 1911;
                int month = tmp.Month;
                int day = tmp.Day;

                string result = "" + year + (month.ToString()).PadLeft(2, '0') + (day.ToString()).PadLeft(2, '0');

                return result;
            }
            else
            {
                return "";
            }
            
            
        }

        /// <summary>
        /// 把文字轉成DateTime
        /// </summary>
        /// <param name="strDate"></param>
        /// <returns></returns>
        public static DateTime? ConvertStringToDateTime(string strDate)
        {
            DateTime dt;

            if(string.IsNullOrEmpty(strDate))
            {
                return null;
            }
            else
            {
                strDate = strDate.Trim();
            }

            if(DateTime.TryParse(strDate, out dt))
            {
                return dt;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 把DB的性別數字轉成輸出的數字
        /// </summary>
        /// <param name="strDbGender"></param>
        /// <returns></returns>
        public static string ConvertDBGenderToOutGender(string strDbGender)
        {
            // 男
            if(strDbGender == "1")
            {
                return "1";
            }
            //女
            else if (strDbGender == "0")
            {
                return "2";
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// 把輸入的性別數字轉成DB的數字
        /// </summary>
        /// <param name="strDbGender"></param>
        /// <returns></returns>
        public static string ConvertOutGenderToDBGender(string strInGender)
        {
            // 男
            if (strInGender == "男")
            {
                return "1";
            }
            // 女
            else if (strInGender == "女")
            {
                return "0";
            }
            else
            {
                return "";
            }
        }

        public static void SetLogData(Log.LogTransfer logTransfer, DAO.StudentFitnessRecord rec)
        {
            // 學年度
            logTransfer.SetLogValue("學年度", rec.SchoolYear.ToString());
            // 測驗日期
            logTransfer.SetLogValue("測驗日期", rec.TestDate.ToShortDateString());
            // 學校類別
            logTransfer.SetLogValue("學校類別", rec.SchoolCategory);
            // 身高
            logTransfer.SetLogValue("身高", rec.Height);
            // 身高常模
            logTransfer.SetLogValue("身高常模", rec.HeightDegree);
            // 體重
            logTransfer.SetLogValue("體重", rec.Weight);
            // 體重常模
            logTransfer.SetLogValue("體重常模", rec.WeightDegree);
            // 坐姿體前彎
            logTransfer.SetLogValue("坐姿體前彎", rec.SitAndReach);
            // 坐姿體前彎常模
            logTransfer.SetLogValue("坐姿體前彎常模", rec.SitAndReachDegree);
            // 立定跳遠
            logTransfer.SetLogValue("立定跳遠", rec.StandingLongJump);
            // 立定跳遠常模
            logTransfer.SetLogValue("立定跳遠常模", rec.StandingLongJumpDegree);
            // 仰臥起坐
            logTransfer.SetLogValue("仰臥起坐", rec.SitUp);
            // 仰臥起坐常模
            logTransfer.SetLogValue("仰臥起坐常模", rec.SitUpDegree);
            // 心肺適能
            logTransfer.SetLogValue("心肺適能", rec.Cardiorespiratory);
            // 心肺適能常模
            logTransfer.SetLogValue("心肺適能常模", rec.CardiorespiratoryDegree);
        }
    }
}
