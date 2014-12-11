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
        public const string KeyFitnessShortCut = "K12.Sports.Fitness.Content.cs.ShortCut";

        public const string KeyFitnessProveSingle = "K12.Sports.Fitness.Report.cs.ProveSingle";

        public const string 體適能常模轉換 = "K12.Sports.FitnessImportExport.ComparisonForm.cs";

        public static bool 體適能常模轉換權限
        {
             get
             {
                  return FISCA.Permission.UserAcl.Current[體適能常模轉換].Executable;
             }
        }

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

        public static bool IsEnableFitnessShortCut
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyFitnessShortCut].Executable;
            }
        }

        public static bool IsEnableFitnessProveSingle
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyFitnessProveSingle].Executable;
            }
        }
    }
}
