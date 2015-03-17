using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using K12.Sports.FitnessImportExport.DAO;

namespace K12.Sports.FitnessImportExport
{
     public class KeyBoStudent
     {
          public KeyBoStudent(DataRow row)
          {
               ID = "" + row["id"];
               Name = "" + row["name"];
               StudentNumber = "" + row["student_number"];
               SeatNo = "" + row["seat_no"];

               DateTime dt;
               if (DateTime.TryParse("" + row["birthdate"], out dt))
               {
                    Birthdate = dt; //生日
               }

               if ("" + row["gender"] == "1")
                    Gender = "男";
               else if ("" + row["gender"] == "0")
                    Gender = "女"; //性別

               //DateTime dt;
               //if (DateTime.TryParse(Birthdate, out dt))
               //{
               //     Age = (DateTime.Today.Year - 1911) - (dt.Year - 1911);
               //}


               ClassID = "" + row["ref_class_id"];
               ClassName = "" + row["class_name"];
          }

          /// <summary>
          /// 體適能資料
          /// </summary>
          public StudentFitnessRecord sfr { get; set; }

          /// <summary>
          /// 體適能LOG資料
          /// </summary>
          public StudentFitnessRecord sfr_Log { get; set; }

          /// <summary>
          /// 姓名
          /// </summary>
          public string Name { get; set; }

          /// <summary>
          /// 學生系統編號
          /// </summary>
          public string ID { get; set; }

          /// <summary>
          /// 生日
          /// </summary>
          public DateTime? Birthdate { get; set; }

          /// <summary>
          /// 年齡
          /// </summary>
          public TestAge Age { get; set; }

          /// <summary>
          /// 性別
          /// </summary>
          public string Gender { get; set; }

          /// <summary>
          /// 學號
          /// </summary>
          public string StudentNumber { get; set; }

          /// <summary>
          /// 座號
          /// </summary>
          public string SeatNo { get; set; }

          /// <summary>
          /// 班級系統編號
          /// </summary>
          public string ClassID { get; set; }

          /// <summary>
          /// 班級名稱
          /// </summary>
          public string ClassName { get; set; }
     }
}
