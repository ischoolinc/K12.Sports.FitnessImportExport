using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Sports.FitnessImportExport
{
    class Permissions
    {
        public const string KeyFitnessExport = "K12.Sports.Fitness.Export.cs";
        public const string KeyFitnessImport = "K12.Sports.Fitness.Import.cs";
        public const string KeyFitnessContent = "K12.Sports.Fitness.Content.cs";
        public static bool IsEnableFitnessExport
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyFitnessExport].Executable;
            }
        }

        public static bool IsEnableFitnessImport
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyFitnessImport].Executable;
            }
        }

        public static bool IsEditableFitnessContent
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyFitnessContent].Editable;
            }
        }

        public static bool IsViewableFitnessContent
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyFitnessContent].Viewable;
            }
        }
    }
}
