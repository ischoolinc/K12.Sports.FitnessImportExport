using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.UDT;


namespace K12.Sports.FitnessImportExport.UDT
{
    /// <summary>
    /// 體適能設定
    /// </summary>
    [TableName("ischool_fitness_config")]
    public class FitnessConfigRecord : ActiveRecord
    {
        /// <summary>
        ///  學年度
        /// </summary>
        [Field(Field = "school_year", Indexed = false)]
        public int SchoolYear { get; set; }

        /// <summary>
        ///  項目名稱
        /// </summary>
        [Field(Field = "item_name", Indexed = false)]
        public string ItemName { get; set; }

        /// <summary>
        ///  設定值
        /// </summary>
        [Field(Field = "item_value", Indexed = false)]
        public string ItemValue { get; set; }
    }
}
