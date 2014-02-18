using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Sports.FitnessImportExport
{
    class Global
    {
        public static readonly string _ModuleName = "體適能";
        public static readonly string _SheetName = "體適能資料";
        public static readonly string[] _ExcelDataTitle =
            { "測驗日期", "學校類別", "年級", "班級名稱", "學號/座號",
                "性別", "身分證字號", "生日", "身高", "體重", "坐姿體前彎", "立定跳遠", "仰臥起坐", "心肺適能" };

        public static readonly string[] _ExcelDataDegreeTitle =
            { "測驗日期", "學校類別", "年級", "班級名稱", "學號/座號",
                "性別", "身分證字號", "生日", "身高", "身高常模", "體重", "體重常模",
                "坐姿體前彎", "坐姿體前彎常模", "立定跳遠", "立定跳遠常模", "仰臥起坐", "仰臥起坐常模", "心肺適能", "心肺適能常模" };

        /// <summary>
        /// 所有學生學號與ID的暫存 key:StudentNumber; value:StudentID
        /// </summary>
        public static Dictionary<string, string> _AllStudentNumberIDTemp = new Dictionary<string, string>();

        /// <summary>
        /// 當有錯誤訊息
        /// </summary>
        public static StringBuilder _ErrorMessageList = new StringBuilder();

        
    }
}
