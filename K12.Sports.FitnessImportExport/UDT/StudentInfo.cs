using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport.DAO
{
    class StudentInfo
    {
        public string Student_ID;
        public string Student_Student_Number;
        public string Student_Gender;
        public string Student_ID_Number;
        public string Student_Birthday;
        public string Class_Grade_Year;
        public string Class_Class_Name;
        public string Student_Name;
        public string Class_Display_Order;

        public StudentInfo(DataRow row)
        {
            int index = 0;
            this.Student_ID = "" + row[index++];
            this.Student_Student_Number = "" + row[index++];
            this.Student_Gender = Utility.ConvertDBGenderToOutGender("" + row[index++]);
            this.Student_ID_Number = "" + row[index++];
            string tmp = "" + row[index++];
            if (!string.IsNullOrEmpty(tmp))  // 假如生日不是空的
            {
                DateTime? dt = Utility.ConvertStringToDateTime(tmp);
                if (dt.HasValue)    // 生日成功轉成DateTime
                    tmp = Utility.ConvertDateTimeToChineseDateTime(dt); // 轉成民國年
            }
            this.Student_Birthday = tmp;
            this.Class_Grade_Year = "" + row[index++];
            this.Class_Class_Name = "" + row[index++];
            this.Student_Name = "" + row[index++];
            this.Class_Display_Order = "" + row[index++];

        }
    }
}
