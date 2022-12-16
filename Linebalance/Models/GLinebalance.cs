using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Linebalance.Models
{
    public class GLinebalance
    {
        public string EQUIPMENT_CODE { get; set; }
        public string CYCLETIME { get; set; }
        public double CAL { get; set; }
        public string CAL_L2 { get; set; }
       

    }
    public class GLinebalance2
    {
        public string CAL_L2 { get; set; }
        public string GROUP_L2 { get; set; }

    }
    public class CalLinebalance
    {
        public string EQUIPMENT_CODE { get; set; }
        public double CYCLETIME { get; set; }
        public double CAL { get; set; }
        public string CAL_L2 { get; set; }


    }




}