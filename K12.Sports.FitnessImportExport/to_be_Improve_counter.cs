using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport
{

    // 2016/7/18 穎驊新增，定義在群集中各項目"待加強"之人數

    class to_be_Improve_counter
    {


        //坐姿體前彎待加強人數
        public decimal SitAndReachDegree_failed_counter { get; set; }


        //坐姿體前彎總人數
        public decimal SitAndReachDegree_total { get; set; }

        //立定跳遠待加強人數
        public decimal StandingLongJumpDegree_failed_counter { get; set; }


        //立定跳遠總人數
        public decimal StandingLongJumpDegree_total { get; set; }


        //仰臥起坐待加強人數
        public decimal SitUpDegree_failed_counter { get; set; }


        //仰臥起坐總人數
        public decimal SitUpDegree_total { get; set; }



        //心肺功能待加強人數
        public decimal CardiorespiratoryDegree_failed_counter { get; set; }

        //心肺功能總人數
        public decimal CardiorespiratoryDegree_total { get; set; }



        //四個項目(坐姿體前彎、立定跳遠、仰臥起坐、心肺適能)全過(在金牌、銀牌、銅牌、中等、待加強五個評等中至少拿中等)人數
        public decimal Four_Item_All_Pass_counter { get; set; }

        //四個項目全過總人數
        public decimal Four_Item_All_Pass_counter_total { get; set; }




    }
    
    


}
