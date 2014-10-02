using K12.Data;
using K12.Sports.FitnessImportExport.DAO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport
{
    class StudentRSRecord
    {
        /// <summary>
        /// 學生本人
        /// </summary>
        public StudentRecord _student { get; set; }

        /// <summary>
        /// 學生體適能記錄
        /// </summary>
        public List<StudentFitnessRecord> _ResultList { get; set; }

        public string 學生入學照片 { get; set; }

        public string 學生畢業照片 { get; set; }

        public StudentRSRecord(StudentRecord student)
        {
            _student = student;
            _ResultList = new List<StudentFitnessRecord>();
        }

        /// <summary>
        /// 設定學生社團參與記錄
        /// </summary>
        public void SetRSR(StudentFitnessRecord rsr)
        {
            _ResultList.Add(rsr);
        }
    }
}
