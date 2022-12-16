using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Linebalance.Models;
using System.Configuration;
using Oracle.ManagedDataAccess.Client;
using Linebalance.Comfunction;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.IO;

using HttpException = System.Web.HttpException;
using Microsoft.Win32;
using System.Web.UI;
using System.Xml;
using System.Xml.Serialization;
using System.Text;

namespace Linebalance.Controllers
{
    public class linebarlanceController : Controller
    {
        static List<R_LOGIN> lLogin = new List<R_LOGIN>();
        static List<R_LOGIN_FAC> lLoginF = new List<R_LOGIN_FAC>();
        static List<R_LOGIN_LINE> lLoginL = new List<R_LOGIN_LINE>();
        static List<R_LOGIN_AREA> lLoginA = new List<R_LOGIN_AREA>();
        static List<LineBalance> Data = new List<LineBalance>();
        static List<OIRate> OIrate = new List<OIRate>();
        static List<Input> InputData = new List<Input>();
        static List<T_MODEL> MO = new List<T_MODEL>();
        static List<GLinebalance> GData = new List<GLinebalance>();
        static List<GLinebalance2> GCalL2 = new List<GLinebalance2>();
        static string plantlogin;
        static string GArea;
        public ActionResult GetloginOA()
        {
            lLogin.Clear();
            GData.Clear();
            //Response.Redirect("http://thbpo-oa-service.delta.corp/OADETSingleSignOn/Check.aspx?URL=http://thbpoprodap-mes.delta.corp/Linebalance/linebarlance/Login");           
            Response.Redirect("http://thbpo-oa-service.delta.corp/OADETSingleSignOn/Check.aspx?URL=https://localhost:44334/linebarlance/Login");
            return RedirectToAction("Login");
        }

        public ActionResult Login()
        {
            lLogin.Clear();
            GData.Clear();
            string user = Request["userempid"];
            string account = Request["userNTAccount"];

            if (account != null)
            {
                R_LOGIN cus = new R_LOGIN();

                cus.emp_no = user;
                cus.emp_account = account;
                lLogin.Add(cus);

                SqlConnection conns = new SqlConnection("Server=THBPOCIMDB; Database=MESPRDDB; User=MESDB; Password=MES12345");
                System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand();
                command.CommandType = System.Data.CommandType.Text;

                command.CommandText = "insert into MES_UTILIZE_LOGIN([USER_ID],[USER_NAME],[DATE_LOGIN]) values('" + user + "','" + account + "','" + DateTime.Now + "')";
                command.Connection = conns;

                conns.Open();
                command.ExecuteNonQuery();
                conns.Close();

            }
            else
            {
                return Redirect("GetloginOA");
            }

            ViewBag.User = user;

            return View(lLogin);
        }
        public static void Getdataline()
        {

 
            LineBalance item = new LineBalance();
            item.EQUIPMENT_CODE = "";
            item.LINE_NAME = "";
            item.STATION_NAME = "";
            item.BEGIN_POINT = "";
            item.END_POINT = "";
            item.CYCLE_TIME = "";
            item.INPUT_QTY = "";
            item.PASS_QTY = "";
            item.FAIL_QTY = "";
            item.WARNING_CNT = "";
            item.RUNNING_TIME = "";
            item.WAITING_TIME = "";
            item.MO_NUMBER = "";
            item.MODEL_NAME = "";
            item.BARCODE = "";
            item.P_ID = "";
            Data.Add(item);
        }
        public ActionResult ZeroLineBalance()
        {

            Data.Clear();
            string plant = lLoginF[0].SCHEMA;
            string factory = InputData[0].FACTORY;
            string line = InputData[0].LINE;
            string area = InputData[0].AREAR;
            string model = InputData[0].MODEL;
            string date1 = InputData[0].DATE1;
            string date2 = InputData[0].DATE2;

            #region DET6-7
            if (plant == "DET_FM" || plant == "DET_CNDC")
            {
                using (OracleConnection conn = ConFunc.GetDBConnection6())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from (SELECT NVL(M.OI_RATE, 20) RATE, T.Line_Name"+
                                  " from "+plant+".R_SCHEDULE_SFCS_T T "+
                                  "LEFT JOIN "+plant+".R_SCHEDULE_SAP_T M " +
                                  "ON T.SCHL_NO = M.SCHL_NO " +
                                  "where LINE_NAME = ('"+line+"')" +
                                  "and t.SFCS_MODEL = '"+model+"'  " +
                                  "and t.start_date BETWEEN to_date('"+date1+"','YYYYMMDD HH24miss') and to_date('"+date2+"','YYYYMMDD HH24miss')  " +
                                  "ORDER BY T.SCHL_NO) " +
                                  "WHERE ROWNUM = 1";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        OIRate OI = new OIRate();
                        OI.RATE = reader["RATE"].ToString();
                        OI.LINE_NAME = reader["LINE_NAME"].ToString();

                        OIrate.Add(OI);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection6())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select b.equipment_code,c.line_name,d.prod_area_desc,c.group_name,c.station_name,c.pqm_name_en,a.begin_point,a.end_point,a.cycle_time,a.input_qty,a.pass_qty,a.fail_qty,a.warning_cnt,a.running_time,a.waiting_time,a.mo_number,a.model_name,a.barcode,a.p_id" +
                                " from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id " +
                                "inner join " + plant + ".c_station_config_t c " +
                                "on a.equipment_id = c.equipment_id " +
                                "inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id" +
                                " where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc ='" + area + "'and a.end_point" +
                                " BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') ";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        LineBalance item = new LineBalance();
                        item.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        item.LINE_NAME = reader["LINE_NAME"].ToString();
                        item.GROUP_NAME = reader["GROUP_NAME"].ToString();
                        item.STATION_NAME = reader["STATION_NAME"].ToString();
                        item.BEGIN_POINT = reader["BEGIN_POINT"].ToString();
                        item.END_POINT = reader["END_POINT"].ToString();
                        item.CYCLE_TIME = reader["CYCLE_TIME"].ToString();
                        item.INPUT_QTY = reader["INPUT_QTY"].ToString();
                        item.PASS_QTY = reader["PASS_QTY"].ToString();
                        item.FAIL_QTY = reader["FAIL_QTY"].ToString();
                        item.WARNING_CNT = reader["WARNING_CNT"].ToString();
                        item.RUNNING_TIME = reader["RUNNING_TIME"].ToString();
                        item.WAITING_TIME = reader["WARNING_CNT"].ToString();
                        item.MO_NUMBER = reader["MO_NUMBER"].ToString();
                        item.MODEL_NAME = reader["MODEL_NAME"].ToString();
                        item.BARCODE = reader["BARCODE"].ToString();
                        item.P_ID = reader["P_ID"].ToString();
                        Data.Add(item);

                    }
                }

                        using (OracleConnection conn = ConFunc.GetDBConnection6())
                        {
                            var _with1 = conn;
                            if (_with1.State == ConnectionState.Open)
                                _with1.Close();
                            _with1.Open();

                            string str = @"select ROUND((CYCLETIME/CAL)/group_qty,2) as cal_linebalance,group_name,EQUIPMENT_CODE from (select b.equipment_code as equipment_code ,d.prod_area_desc,a.model_name,sum(a.CYCLE_TIME)as CYCLETIME, COUNT(*)as CAL,c.group_name as group_name, NVL(e.group_qty, 1)as group_qty from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id inner join " + plant + ".c_station_config_t c on a.equipment_id = c.equipment_id inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id left join " + plant + ".C_MODEL_GROUP_QTY_T e on c.group_name = e.group_name and a.model_name = e.model_name where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc = '" + area + "' and a.end_point BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') group by b.equipment_code,d.prod_area_desc,c.group_name,a.model_name,e.group_qty)";
                            OracleCommand cmd = new OracleCommand(str, conn);
                            OracleDataReader reader;
                            reader = cmd.ExecuteReader();
                            GData.Clear();
                            while (reader.Read())
                            {

                            GLinebalance Gitem = new GLinebalance();
                            Gitem.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                            Gitem.CYCLETIME = reader["cal_linebalance"].ToString();

                            GData.Add(Gitem);

                            }
                        }

               



            }
            #endregion
            #region DET1-5
            else //เงื่อนไขโรง 5
            {

                using (OracleConnection conn = ConFunc.GetDB5Connection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from (SELECT NVL(M.OI_RATE, 20) RATE, T.Line_Name" +
                                  " from " + plant + ".R_SCHEDULE_SFCS_T T " +
                                  "LEFT JOIN " + plant + ".R_SCHEDULE_SAP_T M " +
                                  "ON T.SCHL_NO = M.SCHL_NO " +
                                  "where LINE_NAME = ('" + line + "')" +
                                  "and t.SFCS_MODEL = '" + model + "'  " +
                                  "and t.start_date BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss')  " +
                                  "ORDER BY T.SCHL_NO) " +
                                  "WHERE ROWNUM = 1";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        OIRate OI = new OIRate();
                        OI.RATE = reader["RATE"].ToString();
                        OI.LINE_NAME = reader["LINE_NAME"].ToString();

                        OIrate.Add(OI);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select b.equipment_code,c.line_name,d.prod_area_desc,c.group_name,c.station_name,c.pqm_name_en,a.begin_point,a.end_point,a.cycle_time,a.input_qty,a.pass_qty,a.fail_qty,a.warning_cnt,a.running_time,a.waiting_time,a.mo_number,a.model_name,a.barcode,a.p_id" +
                                " from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id " +
                                "inner join " + plant + ".c_station_config_t c " +
                                "on a.equipment_id = c.equipment_id " +
                                "inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id" +
                                " where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc ='" + area + "'and a.end_point" +
                                " BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') ";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        LineBalance item = new LineBalance();
                        item.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        item.LINE_NAME = reader["LINE_NAME"].ToString();
                        item.GROUP_NAME = reader["GROUP_NAME"].ToString();
                        item.STATION_NAME = reader["STATION_NAME"].ToString();
                        item.BEGIN_POINT = reader["BEGIN_POINT"].ToString();
                        item.END_POINT = reader["END_POINT"].ToString();
                        item.CYCLE_TIME = reader["CYCLE_TIME"].ToString();
                        item.INPUT_QTY = reader["INPUT_QTY"].ToString();
                        item.PASS_QTY = reader["PASS_QTY"].ToString();
                        item.FAIL_QTY = reader["FAIL_QTY"].ToString();
                        item.WARNING_CNT = reader["WARNING_CNT"].ToString();
                        item.RUNNING_TIME = reader["RUNNING_TIME"].ToString();
                        item.WAITING_TIME = reader["WARNING_CNT"].ToString();
                        item.MO_NUMBER = reader["MO_NUMBER"].ToString();
                        item.MODEL_NAME = reader["MODEL_NAME"].ToString();
                        item.BARCODE = reader["BARCODE"].ToString();
                        item.P_ID = reader["P_ID"].ToString();
                        Data.Add(item);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select ROUND((CYCLETIME/CAL)/group_qty,2) as cal_linebalance,group_name, EQUIPMENT_CODE from (select b.equipment_code as equipment_code ,d.prod_area_desc,a.model_name,sum(a.CYCLE_TIME)as CYCLETIME, COUNT(*)as CAL,c.group_name as group_name, NVL(e.group_qty, 1)as group_qty from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id inner join " + plant + ".c_station_config_t c on a.equipment_id = c.equipment_id inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id left join " + plant + ".C_MODEL_GROUP_QTY_T e on c.group_name = e.group_name and a.model_name = e.model_name where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc = '" + area + "' and a.end_point BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') group by b.equipment_code,d.prod_area_desc,c.group_name,a.model_name,e.group_qty)";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    GData.Clear();
                    while (reader.Read())
                    {

                        GLinebalance Gitem = new GLinebalance();
                        Gitem.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        Gitem.CYCLETIME = reader["cal_linebalance"].ToString();

                        GData.Add(Gitem);

                    }
                }
                                           
            }
            #endregion

            #region ExportSave 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo template = new FileInfo(Server.MapPath(@"Z_Template.xlsx"));


            using (var package = new ExcelPackage(template))
            {
                var workbook = package.Workbook;
                var worksheet = package.Workbook.Worksheets["Report"];
                int startRows = 2;


                foreach (var iteme in Data)
                {

                    worksheet.Cells[startRows, 1].Value = iteme.EQUIPMENT_CODE;
                    worksheet.Cells[startRows, 2].Value = iteme.LINE_NAME;
                    worksheet.Cells[startRows, 3].Value = iteme.GROUP_NAME;
                    worksheet.Cells[startRows, 4].Value = iteme.STATION_NAME;
                    worksheet.Cells[startRows, 5].Value = iteme.BEGIN_POINT;
                    worksheet.Cells[startRows, 6].Value = iteme.END_POINT;
                    worksheet.Cells[startRows, 7].Value = iteme.INPUT_QTY;
                    worksheet.Cells[startRows, 8].Value = iteme.PASS_QTY;
                    //worksheet.Cells[startRows, 8].Value = iteme.FAIL_QTY;


                    if (iteme.CYCLE_TIME == "")
                    {
                        worksheet.Cells[startRows, 9].Value = "0.00" + "%";
                    }
                    else
                    {
                        worksheet.Cells[startRows, 9].Value = iteme.CYCLE_TIME + "%";
                    }
                    worksheet.Cells[startRows, 10].Value = iteme.RUNNING_TIME;
                    worksheet.Cells[startRows, 11].Value = iteme.MODEL_NAME;
                    worksheet.Cells[startRows, 12].Value = iteme.FAIL_QTY;
                    worksheet.Cells[startRows, 13].Value = iteme.MO_NUMBER;

                    startRows++;
                }

                //package.SaveAs(new FileInfo(Server.MapPath(@"Utilization Report" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss", System.Globalization.DateTimeFormatInfo.InvariantInfo) + ".xlsx")));
                package.SaveAs(new FileInfo(Server.MapPath(@"Line Balance Report" + ".xlsx")));
            }
            #endregion ExportSave

            ViewData["OI"] = OIrate;
            ViewData["Fac"] = factory;


            return PartialView("BTQry", GData);

        }
        
        public FileResult DownloadFile()
        {

            //Fetch all files in the Folder (Directory).
            string[] filePaths = Directory.GetFiles(Server.MapPath("~/linebarlance/"));

            //Copy File names to Model collection.
            List<FileModel> files = new List<FileModel>();
            foreach (string filePath in filePaths)
            {
                files.Add(new FileModel { FileName = Path.GetFileName(filePath) });
            }

            string fileName = "Line Balance Report" + ".xlsx";

            //Build the File Path.
            string path = Server.MapPath("~/linebarlance/") + fileName;

            //Read the File data into Byte Array.
            byte[] bytes = System.IO.File.ReadAllBytes(path);

            //Send the File to Download.
            return File(bytes, "application/octet-stream", fileName);
        }
        public ActionResult ConLineBalance(List<String> Factory, List<String> Line, List<String> Model, List<String> Date, List<String> Area, List<String> check)
        {
            InputData.Clear();
            Input addInput = new Input();
            addInput.FACTORY = Factory[0];
            addInput.LINE = Line[0];
            addInput.AREAR = Area[0];
            addInput.MODEL = Model[0];
            string date = Date[0];
            string[] strcut = Date[0].Split(" ".ToCharArray());
            addInput.DATE1 = Convert.ToDateTime(strcut[0].ToString()).ToString("yyyyMMdd");
            addInput.DATE2 = Convert.ToDateTime(strcut[2].ToString()).ToString("yyyyMMdd");
            InputData.Add(addInput);
            string factory = InputData[0].FACTORY;
            if (InputData[0].FACTORY == "FMBG" || InputData[0].FACTORY == "CNDC")
            {
                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from SFCS.C_FACTORY_AREA_T where FACTORY = '" + factory + "'";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginF.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_FAC item = new R_LOGIN_FAC();
                        item.SCHEMA = reader["SCHEMA"].ToString();
                        lLoginF.Add(item);

                    }
                }
            }
            else 
            {
                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from SFCS.C_FACTORY_AREA_T where FACTORY = '" + factory + "'";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginF.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_FAC item = new R_LOGIN_FAC();
                        item.SCHEMA = reader["SCHEMA"].ToString();
                        lLoginF.Add(item);

                    }
                }

            }
            string cCheck = check[0];
            if (cCheck == "True")
            {
                return RedirectToAction("ZeroLineBalance");
            }
            return RedirectToAction("NonZeroLineBalance");
        }

        public ActionResult NonZeroLineBalance()
        {

            Data.Clear();
            string plant = lLoginF[0].SCHEMA;
            string factory = InputData[0].FACTORY;
            string line = InputData[0].LINE;
            string area = InputData[0].AREAR;
            string model = InputData[0].MODEL;
            string date1 = InputData[0].DATE1;
            string date2 = InputData[0].DATE2;

            #region DET6-7
            if (plant == "DET_FM" || plant == "DET_CNDC")
            {
                using (OracleConnection conn = ConFunc.GetDBConnection6())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from (SELECT NVL(M.OI_RATE, 20) RATE, T.Line_Name" +
                                  " from " + plant + ".R_SCHEDULE_SFCS_T T " +
                                  "LEFT JOIN " + plant + ".R_SCHEDULE_SAP_T M " +
                                  "ON T.SCHL_NO = M.SCHL_NO " +
                                  "where LINE_NAME = ('" + line + "')" +
                                  "and t.SFCS_MODEL = '" + model + "'  " +
                                  "and t.start_date BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss')  " +
                                  "ORDER BY T.SCHL_NO) " +
                                  "WHERE ROWNUM = 1";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        OIRate OI = new OIRate();
                        OI.RATE = reader["RATE"].ToString();
                        OI.LINE_NAME = reader["LINE_NAME"].ToString();

                        OIrate.Add(OI);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection6())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select b.equipment_code,c.line_name,d.prod_area_desc,c.group_name,c.station_name,c.pqm_name_en,a.begin_point,a.end_point,a.cycle_time,a.input_qty,a.pass_qty,a.fail_qty,a.warning_cnt,a.running_time,a.waiting_time,a.mo_number,a.model_name,a.barcode,a.p_id" +
                                " from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id " +
                                "inner join " + plant + ".c_station_config_t c " +
                                "on a.equipment_id = c.equipment_id " +
                                "inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id" +
                                " where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc ='" + area + "'and a.end_point" +
                                " BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') ";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        LineBalance item = new LineBalance();
                        item.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        item.LINE_NAME = reader["LINE_NAME"].ToString();
                        item.GROUP_NAME = reader["GROUP_NAME"].ToString();
                        item.STATION_NAME = reader["STATION_NAME"].ToString();
                        item.BEGIN_POINT = reader["BEGIN_POINT"].ToString();
                        item.END_POINT = reader["END_POINT"].ToString();
                        item.CYCLE_TIME = reader["CYCLE_TIME"].ToString();
                        item.INPUT_QTY = reader["INPUT_QTY"].ToString();
                        item.PASS_QTY = reader["PASS_QTY"].ToString();
                        item.FAIL_QTY = reader["FAIL_QTY"].ToString();
                        item.WARNING_CNT = reader["WARNING_CNT"].ToString();
                        item.RUNNING_TIME = reader["RUNNING_TIME"].ToString();
                        item.WAITING_TIME = reader["WARNING_CNT"].ToString();
                        item.MO_NUMBER = reader["MO_NUMBER"].ToString();
                        item.MODEL_NAME = reader["MODEL_NAME"].ToString();
                        item.BARCODE = reader["BARCODE"].ToString();
                        item.P_ID = reader["P_ID"].ToString();
                        Data.Add(item);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection6())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select ROUND((CYCLETIME/CAL)/group_qty,2) as cal_linebalance,group_name,EQUIPMENT_CODE from (select b.equipment_code as equipment_code ,d.prod_area_desc,a.model_name,sum(a.CYCLE_TIME)as CYCLETIME, COUNT(*)as CAL,c.group_name as group_name, NVL(e.group_qty, 1)as group_qty from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id inner join " + plant + ".c_station_config_t c on a.equipment_id = c.equipment_id inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id left join " + plant + ".C_MODEL_GROUP_QTY_T e on c.group_name = e.group_name and a.model_name = e.model_name where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc = '" + area + "' and a.end_point BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') group by b.equipment_code,d.prod_area_desc,c.group_name,a.model_name,e.group_qty)";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    GData.Clear();
                    while (reader.Read())
                    {

                        GLinebalance Gitem = new GLinebalance();
                        Gitem.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        Gitem.CYCLETIME = reader["cal_linebalance"].ToString();

                        GData.Add(Gitem);

                    }
                }





            }
            #endregion
            #region DET1-5
            else //เงื่อนไขโรง 5
            {

                using (OracleConnection conn = ConFunc.GetDB5Connection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from (SELECT NVL(M.OI_RATE, 20) RATE, T.Line_Name" +
                                  " from " + plant + ".R_SCHEDULE_SFCS_T T " +
                                  "LEFT JOIN " + plant + ".R_SCHEDULE_SAP_T M " +
                                  "ON T.SCHL_NO = M.SCHL_NO " +
                                  "where LINE_NAME = ('" + line + "')" +
                                  "and t.SFCS_MODEL = '" + model + "'  " +
                                  "and t.start_date BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss')  " +
                                  "ORDER BY T.SCHL_NO) " +
                                  "WHERE ROWNUM = 1";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        OIRate OI = new OIRate();
                        OI.RATE = reader["RATE"].ToString();
                        OI.LINE_NAME = reader["LINE_NAME"].ToString();

                        OIrate.Add(OI);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select b.equipment_code,c.line_name,d.prod_area_desc,c.group_name,c.station_name,c.pqm_name_en,a.begin_point,a.end_point,a.cycle_time,a.input_qty,a.pass_qty,a.fail_qty,a.warning_cnt,a.running_time,a.waiting_time,a.mo_number,a.model_name,a.barcode,a.p_id" +
                                " from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id " +
                                "inner join " + plant + ".c_station_config_t c " +
                                "on a.equipment_id = c.equipment_id " +
                                "inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id" +
                                " where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc ='" + area + "'and a.end_point" +
                                " BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') ";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {

                        LineBalance item = new LineBalance();
                        item.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        item.LINE_NAME = reader["LINE_NAME"].ToString();
                        item.GROUP_NAME = reader["GROUP_NAME"].ToString();
                        item.STATION_NAME = reader["STATION_NAME"].ToString();
                        item.BEGIN_POINT = reader["BEGIN_POINT"].ToString();
                        item.END_POINT = reader["END_POINT"].ToString();
                        item.CYCLE_TIME = reader["CYCLE_TIME"].ToString();
                        item.INPUT_QTY = reader["INPUT_QTY"].ToString();
                        item.PASS_QTY = reader["PASS_QTY"].ToString();
                        item.FAIL_QTY = reader["FAIL_QTY"].ToString();
                        item.WARNING_CNT = reader["WARNING_CNT"].ToString();
                        item.RUNNING_TIME = reader["RUNNING_TIME"].ToString();
                        item.WAITING_TIME = reader["WARNING_CNT"].ToString();
                        item.MO_NUMBER = reader["MO_NUMBER"].ToString();
                        item.MODEL_NAME = reader["MODEL_NAME"].ToString();
                        item.BARCODE = reader["BARCODE"].ToString();
                        item.P_ID = reader["P_ID"].ToString();
                        Data.Add(item);

                    }
                }

                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select ROUND((CYCLETIME/CAL)/group_qty,2) as cal_linebalance,group_name, EQUIPMENT_CODE from (select b.equipment_code as equipment_code ,d.prod_area_desc,a.model_name,sum(a.CYCLE_TIME)as CYCLETIME, COUNT(*)as CAL,c.group_name as group_name, NVL(e.group_qty, 1)as group_qty from " + plant + ".r_equipment_pub_param_record_t a inner join " + plant + ".c_equipment_basic_t b on a.equipment_id = b.equipment_id inner join " + plant + ".c_station_config_t c on a.equipment_id = c.equipment_id inner JOIN " + plant + ".C_PROD_AREA_T d on b.prod_area_id = d.prod_area_id left join " + plant + ".C_MODEL_GROUP_QTY_T e on c.group_name = e.group_name and a.model_name = e.model_name where c.line_name = '" + line + "' and a.model_name = '" + model + "' and d.prod_area_desc = '" + area + "' and a.end_point BETWEEN to_date('" + date1 + "','YYYYMMDD HH24miss') and to_date('" + date2 + "','YYYYMMDD HH24miss') group by b.equipment_code,d.prod_area_desc,c.group_name,a.model_name,e.group_qty)";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    GData.Clear();
                    while (reader.Read())
                    {

                        GLinebalance Gitem = new GLinebalance();
                        Gitem.EQUIPMENT_CODE = reader["EQUIPMENT_CODE"].ToString();
                        Gitem.CYCLETIME = reader["cal_linebalance"].ToString();

                        GData.Add(Gitem);

                    }
                }

            }
            #endregion

            #region ExportSave 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo template = new FileInfo(Server.MapPath(@"Z_Template.xlsx"));


            using (var package = new ExcelPackage(template))
            {
                var workbook = package.Workbook;
                var worksheet = package.Workbook.Worksheets["Report"];
                int startRows = 2;


                foreach (var iteme in Data)
                {

                    worksheet.Cells[startRows, 1].Value = iteme.EQUIPMENT_CODE;
                    worksheet.Cells[startRows, 2].Value = iteme.LINE_NAME;
                    worksheet.Cells[startRows, 3].Value = iteme.GROUP_NAME;
                    worksheet.Cells[startRows, 4].Value = iteme.STATION_NAME;
                    worksheet.Cells[startRows, 5].Value = iteme.BEGIN_POINT;
                    worksheet.Cells[startRows, 6].Value = iteme.END_POINT;
                    worksheet.Cells[startRows, 7].Value = iteme.INPUT_QTY;
                    worksheet.Cells[startRows, 8].Value = iteme.PASS_QTY;
                    //worksheet.Cells[startRows, 8].Value = iteme.FAIL_QTY;


                    if (iteme.CYCLE_TIME == "")
                    {
                        worksheet.Cells[startRows, 9].Value = "0.00" + "%";
                    }
                    else
                    {
                        worksheet.Cells[startRows, 9].Value = iteme.CYCLE_TIME + "%";
                    }
                    worksheet.Cells[startRows, 10].Value = iteme.RUNNING_TIME;
                    worksheet.Cells[startRows, 11].Value = iteme.MODEL_NAME;
                    worksheet.Cells[startRows, 12].Value = iteme.FAIL_QTY;
                    worksheet.Cells[startRows, 13].Value = iteme.MO_NUMBER;

                    startRows++;
                }

                //package.SaveAs(new FileInfo(Server.MapPath(@"Utilization Report" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss", System.Globalization.DateTimeFormatInfo.InvariantInfo) + ".xlsx")));
                package.SaveAs(new FileInfo(Server.MapPath(@"Line Balance Report" + ".xlsx")));
            }
            #endregion ExportSave

            ViewData["OI"] = OIrate;
            ViewData["Fac"] = plant;


            return PartialView("BTQry", GData);
        }

        [HttpPost]
        public ActionResult LineBanace(String plant)
        {
            string plantlogin = plant;

            GLinebalance Gitem = new GLinebalance();
            Gitem.EQUIPMENT_CODE = "";
            Gitem.CYCLETIME = "0.0";
            GData.Add(Gitem);

            if (plant == "DET_FM" || plant == "DET_CNDC")
            {
                using (OracleConnection conn5 = ConFunc.GetDB6Connection())
                {
                    var _with1 = conn5;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from SFCS.C_FACTORY_AREA_T where SCHEMA = '" + plant + "'";
                    OracleCommand cmd = new OracleCommand(str, conn5);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginF.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_FAC item = new R_LOGIN_FAC();
                        item.FACTORY = reader["FACTORY"].ToString();
                        item.SCHEMA = reader["SCHEMA"].ToString();
                        lLoginF.Add(item);

                    }

                }
                using (OracleConnection conn5 = ConFunc.GetDB6Connection())
                {
                    var _with1 = conn5;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from " + plant + ".C_MODEL_DESC_T";
                    OracleCommand cmd = new OracleCommand(str, conn5);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    MO.Clear();
                    while (reader.Read())
                    {

                        T_MODEL item = new T_MODEL();
                        item.MODEL_NAME = reader["MODEL_NAME"].ToString();
                        MO.Add(item);

                    }
                }
                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select prod_area_desc from SFCS.C_PROD_AREA_T where Factory = '"+ lLoginF[0].FACTORY + "'";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginA.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_AREA item = new R_LOGIN_AREA();
                        item.PROD_AREA_DESC = reader["PROD_AREA_DESC"].ToString();
                        lLoginA.Add(item);

                    }


                }
                //using (OracleConnection conn = ConFunc.GetDBConnection6())
                //{
                //    var _with1 = conn;
                //    if (_with1.State == ConnectionState.Open)
                //        _with1.Close();
                //    _with1.Open();

                //    string str = @"select a.line_name from " + plant + ".C_LINE_DESC_T a where a.prod_area_id = '"+ lLoginA[0].PROD_AREA_DESC +"'";
                //    OracleCommand cmd = new OracleCommand(str, conn);
                //    OracleDataReader reader;
                //    reader = cmd.ExecuteReader();
                //    lLoginL.Clear();
                //    while (reader.Read())
                //    {

                //        R_LOGIN_LINE item = new R_LOGIN_LINE();
                //        item.LINE_NAME = reader["LINE_NAME"].ToString();
                //        lLoginL.Add(item);

                //    }


                //}
            }
            else
            {
                using (OracleConnection conn5 = ConFunc.GetDB5Connection())
                {
                    var _with1 = conn5;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from SFCS.C_FACTORY_AREA_T where SCHEMA = '" + plant + "'";
                    OracleCommand cmd = new OracleCommand(str, conn5);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginF.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_FAC item = new R_LOGIN_FAC();
                        item.FACTORY = reader["FACTORY"].ToString();
                        item.SCHEMA = reader["SCHEMA"].ToString();
                        lLoginF.Add(item);

                    }

                }
                using (OracleConnection conn5 = ConFunc.GetDB5Connection())
                {
                    var _with1 = conn5;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select * from " + plant + ".C_MODEL_DESC_T";
                    OracleCommand cmd = new OracleCommand(str, conn5);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    MO.Clear();
                    while (reader.Read())
                    {

                        T_MODEL item = new T_MODEL();
                        item.MODEL_NAME = reader["MODEL_NAME"].ToString();
                        MO.Add(item);

                    }
                }
                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select b.prod_area_desc from " + plant + ".C_LINE_DESC_T a inner join  SFCS.C_PROD_AREA_T b on a.prod_area_id = b.prod_area_id group by b.prod_area_desc";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginA.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_AREA item = new R_LOGIN_AREA();
                        item.PROD_AREA_DESC = reader["PROD_AREA_DESC"].ToString();
                        lLoginA.Add(item);

                    }


                }
                using (OracleConnection conn = ConFunc.GetDBConnection())
                {
                    var _with1 = conn;
                    if (_with1.State == ConnectionState.Open)
                        _with1.Close();
                    _with1.Open();

                    string str = @"select a.line_name from "+plant+".C_LINE_DESC_T a inner join  SFCS.C_PROD_AREA_T b on a.prod_area_id = b.prod_area_id";
                    OracleCommand cmd = new OracleCommand(str, conn);
                    OracleDataReader reader;
                    reader = cmd.ExecuteReader();
                    lLoginL.Clear();
                    while (reader.Read())
                    {

                        R_LOGIN_LINE item = new R_LOGIN_LINE();
                        item.LINE_NAME = reader["LINE_NAME"].ToString();
                        lLoginL.Add(item);

                    }


                }

            }

            




            ViewData["Fac"] = lLoginF;
            ViewData["Area"] = lLoginA;
            ViewData["Line"] = lLoginL;
            ViewData["Model"] = MO;

            return PartialView("LineBanace", GData);
        }


    }
}