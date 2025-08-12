using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Sports.FitnessImportExport
{
    class Global
    {
        /// <summary>
        /// 資料項目名稱
        /// </summary>
        public static readonly string _ModuleName = "體適能";

        /// <summary>
        /// Sheet名稱
        /// </summary>
        public static readonly string _SheetName = "體適能資料";

        /// <summary>
        /// 不匯出常模
        /// </summary>
        public static readonly string[] _ExcelDataTitle =
            { "*測驗日期（民國年月份日期）",
            "學校類別",
            "*年級",
            "*班級名稱（需與系統設定一致）",
            "*學號",
            "*性別（男：1、女：2）",
            "*身分證字號",
            "*生日（民國年月份日期）",
            "*身高（公分）",
            "*體重（公斤）",
            "坐姿體前彎（公分）",
            "立定跳遠（公分）",
            "仰臥捲腹（次數）",
            "800/1600公尺跑走（分.秒）",
            "漸速耐力跑（總趟數）" };

        /// <summary>
        /// 是否匯出常模
        /// </summary>
        public static readonly string[] _ExcelDataDegreeTitle =
            { "*測驗日期（民國年月份日期）",
            "學校類別",
            "*年級",
            "*班級名稱（需與系統設定一致）",
            "*學號",
            "*性別（男：1、女：2）",
            "*身分證字號",
            "*生日（民國年月份日期）",
            "*身高（公分）",
            "*體重（公斤）",
            "坐姿體前彎（公分）", 
            "立定跳遠（公分）", 
            "仰臥捲腹（次數）", 
            "800/1600公尺跑走（分.秒）",
            "漸速耐力跑（總趟數）",
            "坐姿體前彎常模",
            "立定跳遠常模",
            "仰臥捲腹常模",
            "800/1600公尺跑走常模",
            "漸速耐力跑常模" };


        /// <summary>
        /// 所有學生學號與ID的暫存 key:StudentNumber; value:StudentID
        /// </summary>
        public static Dictionary<string, string> _AllStudentNumberIDTemp = new Dictionary<string, string>();

        /// <summary>
        /// 所有學生身分證號與ID的暫存 key:ID_Number; value:StudentID
        /// </summary>
        public static Dictionary<string, string> _AllStudentIDNumberIDTemp = new Dictionary<string, string>();

        /// <summary>
        /// 當有錯誤訊息
        /// </summary>
        public static StringBuilder _ErrorMessageList = new StringBuilder();


    }
}
