using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.UDT;
using FISCA.DSAUtil;

namespace K12.Sports.FitnessImportExport.DAO
{
    /// <summary>
    /// 處理 UDT 資料
    /// </summary>
    class StudentFitness
    {
        /// <summary>
        /// 建立使用到的 UDT Table：主要檢查資料庫有沒有建立UDT，沒有建自動建立。
        /// </summary>
        public static void CreateFitnessUDTTable()
        {
            FISCA.UDT.SchemaManager Manager = new SchemaManager(new DSConnection(FISCA.Authentication.DSAServices.DefaultDataSource));

            // 學生體適能
            Manager.SyncSchema(new StudentFitnessRecord());
        }

        /// <summary>
        /// 新增學生體適能
        /// </summary>
        /// <param name="DataList"></param>
        public static void InsertByRecordList(List<StudentFitnessRecord> DataList)
        {
            if (DataList.Count > 0)
            {
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.InsertValues(DataList);
            }
        }

        /// <summary>
        /// 新增學生體適能
        /// </summary>
        /// <param name="DataList"></param>
        public static void InsertByRecord(StudentFitnessRecord rec)
        {
            if (rec != null)
            {
                List<StudentFitnessRecord> insertList = new List<StudentFitnessRecord>();
                insertList.Add(rec);
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.InsertValues(insertList);
            }
        }

        /// <summary>
        /// 更新學生體適能
        /// </summary>
        /// <param name="DataList"></param>
        public static void UpdateByRecordList(List<StudentFitnessRecord> DataList)
        {
            if (DataList.Count > 0)
            {
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.UpdateValues(DataList);
            }
        }

        /// <summary>
        /// 更新學生體適能
        /// </summary>
        /// <param name="DataList"></param>
        public static void UpdateByRecord(StudentFitnessRecord rec)
        {
            if (rec != null)
            {
                List<StudentFitnessRecord> updateList = new List<StudentFitnessRecord>();
                updateList.Add(rec);
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.UpdateValues(updateList);
            }
        }

        /// <summary>
        /// 依學生ID 取得學生體適能資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<StudentFitnessRecord> SelectByStudentIDList(List<string> StudentIDList)
        {
            List<StudentFitnessRecord> dataList = new List<StudentFitnessRecord>();
            if (StudentIDList.Count > 0)
            {
                AccessHelper accessHelper = new AccessHelper();
                // 當有 Where 條件寫法
                string query = "ref_student_id in ('" + string.Join("','", StudentIDList.ToArray()) + "')";
                dataList = accessHelper.Select<StudentFitnessRecord>(query);
            }
            return dataList;
        }

        /// <summary>
        /// 依學生ID 取得學生體適能資料
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public static List<StudentFitnessRecord> SelectByStudentID(string StudentID)
        {
            List<StudentFitnessRecord> dataList = new List<StudentFitnessRecord>();
            if(string.IsNullOrEmpty(StudentID))
                return dataList;

            AccessHelper accessHelper = new AccessHelper();
            // 當有 Where 條件寫法
            string query = "ref_student_id = '" + StudentID + "'";
            dataList = accessHelper.Select<StudentFitnessRecord>(query);

            if(dataList == null)
                return new List<StudentFitnessRecord>();

            return dataList;
        }

        /// <summary>
        /// 依學生ID 以及學年度 取得學生體適能資料
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public static List<StudentFitnessRecord> SelectByStudentIDAndSchoolYear(string StudentID, int SchoolYear)
        {
            List<StudentFitnessRecord> dataList = new List<StudentFitnessRecord>();
            if (string.IsNullOrEmpty(StudentID))
                return dataList;

            AccessHelper accessHelper = new AccessHelper();
            // 當有 Where 條件寫法
            string query = "ref_student_id = '" + StudentID + "' and school_year=" + SchoolYear;
            dataList = accessHelper.Select<StudentFitnessRecord>(query);

            if (dataList == null)
                return new List<StudentFitnessRecord>();

            return dataList;
        }

        /// <summary>
        /// 依學生ID 以及學年度 取得學生體適能資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <returns></returns>
        public static List<StudentFitnessRecord> SelectByStudentIDListAndSchoolYear(List<string> StudentIDList, int SchoolYear)
        {
            List<StudentFitnessRecord> dataList = new List<StudentFitnessRecord>();
            if ((StudentIDList.Count > 0) && (SchoolYear > 0))
            {
                AccessHelper accessHelper = new AccessHelper();
                // 當有 Where 條件寫法
                string query = "ref_student_id in ('" + string.Join("','", StudentIDList.ToArray()) + "') and school_year=" + SchoolYear;
                dataList = accessHelper.Select<StudentFitnessRecord>(query);
            }
            return dataList;
        }

        /// <summary>
        /// 依學生ID 以及學年度 取得學生體適能資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <returns></returns>
        public static List<StudentFitnessRecord> SelectByStudentIDListAndSchoolYear(List<string> StudentIDList, List<string> SchoolYearList)
        {
            List<StudentFitnessRecord> dataList = new List<StudentFitnessRecord>();
            if ((StudentIDList.Count > 0) && (SchoolYearList.Count > 0))
            {
                AccessHelper accessHelper = new AccessHelper();
                // 當有 Where 條件寫法
                string query = "ref_student_id in ('" + string.Join("','", StudentIDList.ToArray()) + "') and school_year in (" + string.Join(",", SchoolYearList.ToArray()) + ")";
                dataList = accessHelper.Select<StudentFitnessRecord>(query);
            }
            return dataList;
        }

        /// <summary>
        /// 刪除學生體適能資料
        /// </summary>
        /// <param name="DataList"></param>
        public static void DeleteByRecordList(List<StudentFitnessRecord> DataList)
        {
            if (DataList.Count > 0)
            {
                foreach (StudentFitnessRecord data in DataList)
                {
                    // Deleted 設成 true 才會真刪除
                    data.Deleted = true;
                }
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.DeletedValues(DataList);
            }
        }

        /// <summary>
        /// 刪除學生體適能資料
        /// </summary>
        /// <param name="rec"></param>
        public static void DeleteByRecord(StudentFitnessRecord rec)
        {
            if (rec != null)
            {
                // Deleted 設成 true 才會真刪除
                rec.Deleted = true;

                List<StudentFitnessRecord> recList = new List<StudentFitnessRecord>();
                recList.Add(rec);

                AccessHelper accessHelper = new AccessHelper();
                accessHelper.DeletedValues(recList);
            }
        }
    }
}
