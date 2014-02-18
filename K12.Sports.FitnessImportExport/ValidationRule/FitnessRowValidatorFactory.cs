using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Campus.DocumentValidator;

namespace K12.Sports.FitnessImportExport.ValidationRule
{
    class FitnessRowValidatorFactory : IRowValidatorFactory
    {
        #region IRowValidatorFactory 成員

        IRowVaildator IRowValidatorFactory.CreateRowValidator(string typeName, System.Xml.XmlElement validatorDescription)
        {
            switch (typeName.ToUpper())
            {
                case "K12SPORTFITNESSCHECKSTUDENTNUMBER":
                    return new RowValidator.StudCheckStudentNumberVal();
                default:
                    return null;
            }
        }

        #endregion
    }
}
