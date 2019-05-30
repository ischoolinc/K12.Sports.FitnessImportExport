using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport
{
     static public class tool
     {
          static public FISCA.UDT.AccessHelper _A = new FISCA.UDT.AccessHelper();

          static public FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();

          /// <summary>
          /// 取得狀態為 一般,輟學 學生清單
          /// </summary>
          public static Dictionary<string, KeyBoStudent> GetStudentList(List<string> StudentIDList)
          {
               Dictionary<string, KeyBoStudent> dic = new Dictionary<string, KeyBoStudent>();

               StringBuilder sb = new StringBuilder();
               sb.Append("select student.id,student.name,student.student_number,student.seat_no,student.birthdate,student.gender,student.ref_class_id,class.class_name from student ");
               sb.Append("left join class on student.ref_class_id=class.id ");
               sb.Append(string.Format("where student.id in ('{0}')", string.Join("','", StudentIDList)));

               DataTable dt = tool._Q.Select(sb.ToString());
               foreach (DataRow row in dt.Rows)
               {
                    KeyBoStudent stud = new KeyBoStudent(row);

                    if (!dic.ContainsKey(stud.ID))
                    {
                         dic.Add(stud.ID, stud);
                    }
               }

               return dic;
          }
     }
}
