using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport.DAO
{
    /// <summary>
    /// 使用 FISCA.Data Query
    /// </summary>
    class FDQuery
    {
        /// <summary>
        /// 取得所有學生學號 key:StudentNumber ; value:StudentID
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetAllStudenNumberDict()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select student.student_number,student.id from student where student.status='1';";
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString();
                string value = dr[1].ToString();
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, value);
            }

            return retVal;
        }

        /// <summary>
        /// 取得所有學生身分證號 key:ID_Number ; value:StudentID
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetAllIDNumberDict()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select student.id_number,student.id from student where student.status='1'";
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString(); //身分證號
                string value = dr[1].ToString();
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, value);
            }

            return retVal;
        }

        /// <summary>
        /// 依學號取得學生ID StudentNumber,id
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetStudenIdDictByStudentNumber(List<string> StudNumList)
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            StringBuilder sbQry = new StringBuilder();
            sbQry.Append("'"); sbQry.Append(string.Join("','", StudNumList.ToArray())); sbQry.Append("'");
            string strSQL = "select student.student_number, student.id from student where student_number in(" + sbQry.ToString() + ") and student.status='1';";
            DataTable dt = qh.Select(strSQL);

            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString();
                string value = dr[1].ToString();
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, value);
            }

            return retVal;
        }

        /// <summary>
        /// 依身分證號取得學生ID
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetStudenByIDNumber(List<string> StudIDNumList)
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select student.id_number, student.id from student where id_number in('" + string.Join("','", StudIDNumList) + "') and student.status='1'";
            DataTable dt = qh.Select(strSQL);

            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString();
                string value = dr[1].ToString();
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, value);
            }

            return retVal;
        }


        public static List<StudentInfo> GetStudnetInfoByIDList(List<string> StudIdList)
        {
            List<StudentInfo> result = new List<StudentInfo>();
            QueryHelper qh = new QueryHelper();
            StringBuilder sbQry = new StringBuilder();
            sbQry.Append("'"); sbQry.Append(string.Join("','", StudIdList.ToArray())); sbQry.Append("'");
            string strSQL = "select student.id," +
                                    "student.student_number," +
                                    "student.gender," +
                                    "student.id_number," +
                                    "student.birthdate," +
                                    "class.grade_year," +
                                    "class.class_name," +
                                "student.name," +       // for sort
                                "class.display_order" + // for sort
                            " from student" +
                            " left join class on student.ref_class_id = class.id" +
                            " where student.id in(" + sbQry.ToString() + ") and student.status='1';";
            DataTable dt = qh.Select(strSQL);

            foreach (DataRow dr in dt.Rows)
            {
                StudentInfo studentInfo = new StudentInfo(dr);
                result.Add(studentInfo);
            }

            return result;
        }
    }
}
