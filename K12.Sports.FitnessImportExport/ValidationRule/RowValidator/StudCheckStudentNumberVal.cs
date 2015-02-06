using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Campus.DocumentValidator;

namespace K12.Sports.FitnessImportExport.ValidationRule.RowValidator
{
    class StudCheckStudentNumberVal : IRowVaildator
    {
        
        public string Correct(IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(IRowStream row)
        {
            bool retVal = false;
            string key = Utility.GetIRowValueString(row, "身分證字號");

            if (Global._AllStudentIDNumberIDTemp.ContainsKey(key))
                retVal = true;

            return retVal;
        }

    }
}
