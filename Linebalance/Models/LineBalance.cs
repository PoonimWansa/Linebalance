using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Linebalance.Models
{
    public class LineBalance
    {

        public string EQUIPMENT_CODE { get; set; }
        public string GROUP_NAME { get; set; }
        public string LINE_NAME { get; set; }
        public string STATION_NAME { get; set; }
        public string BEGIN_POINT { get; set; }
        public string END_POINT { get; set; }
        public string CYCLE_TIME { get; set; }
        public string INPUT_QTY { get; set; }
        public string PASS_QTY { get; set; }
        public string FAIL_QTY { get; set; }
        public string WARNING_CNT { get; set; }
        public string RUNNING_TIME { get; set; }
        public string WAITING_TIME { get; set; }
        public string MO_NUMBER { get; set; }
        public string MODEL_NAME { get; set; }
        public string BARCODE { get; set; }
        public string P_ID { get; set; }

    }
    public class OIRate
    {

        public string RATE { get; set; }
        public string LINE_NAME { get; set; }
        

    }
    public class FileModel
    {
        public string FileName { get; set; }
    }
}