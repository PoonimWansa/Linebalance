using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Linebalance.Comfunction
{
    public class ConFunc
    {
        public static OracleConnection GetDB5Connection()
        {

            //Console.WriteLine("Getting Connection ...");

            string host = "THPUBMES-SCAN";
            int port = 1521;
            string sid = "THPUBMES";
            string user = "MESAP03";
            string password = "Delta12345";

            // 'Connection string' to connect directly to Oracle.
            string connString = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = "
                 + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = "
                 + sid + ")));Password=" + password + ";User ID=" + user;

            return new OracleConnection(connString);
        }
        public static OracleConnection GetDB6Connection()
        {

            // Console.WriteLine("Getting Connection6 ...");

            string host = "172.19.249.3";
            int port = 1521;
            string sid = "DETBCWG";
            string user = "MESAP03";
            string password = "Delta12345";

            // 'Connection string' to connect directly to Oracle.
            string connString = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = "
                 + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = "
                 + sid + ")));Password=" + password + ";User ID=" + user;

            return new OracleConnection(connString);
        }


        public static OracleConnection GetDBConnection()
        {

            //Console.WriteLine("Getting Connection ...");

            string host = "THBPODSMRACDB";
            int port = 1521;
            string sid = "THSFDB";
            string user = "PQM_QUERY";
            string password = "tmudk$o0";

            // 'Connection string' to connect directly to Oracle.
            string connString = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = "
                 + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = "
                 + sid + ")));Password=" + password + ";User ID=" + user;

            //OracleConnection conn = new OracleConnection();

            //conn.ConnectionString = connString;

            //return conn;
            return new OracleConnection(connString);
        }
        public static OracleConnection GetDBConnection6()
        {

            // Console.WriteLine("Getting Connection6 ...");

            string host = "thbpomesrpt00";
            int port = 1521;
            string sid = "THRPTDB";
            string user = "PQM_QUERY";
            string password = "tmudk$o0";

            // 'Connection string' to connect directly to Oracle.
            string connString = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = "
                 + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = "
                 + sid + ")));Password=" + password + ";User ID=" + user;

            return new OracleConnection(connString);
        }
        

    }
}