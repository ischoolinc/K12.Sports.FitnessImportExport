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
            { "測驗日期", "學校類別", "年級", "班級名稱", "學號/座號",
                "性別", "身分證字號", "生日", "身高", "體重", "坐姿體前彎", "立定跳遠", "仰臥捲腹", "心肺耐力","漸速耐力跑" };

        /// <summary>
        /// 是否匯出常模
        /// </summary>
        public static readonly string[] _ExcelDataDegreeTitle =
            { "測驗日期", "學校類別", "年級", "班級名稱", "學號/座號",
                "性別", "身分證字號", "生日", "身高", "體重", "坐姿體前彎", "坐姿體前彎常模", "立定跳遠", "立定跳遠常模", "仰臥捲腹", "仰臥捲腹常模", "心肺耐力", "心肺耐力常模","漸速耐力跑","漸速耐力跑常模" };


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
