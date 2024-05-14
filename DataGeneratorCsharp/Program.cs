using System;
using System.Data.SqlClient;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
//https://github.com/ExcelDataReader/ExcelDataReader
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace DatageneratorCsharp
{
    internal class Program
    {
        public class HeartRateData
        {
            public int SensorId;
            public int HeartRateBPM;
            public DateTime TimeStamp;
            public string Label;
            public string Condition;
        }

        static void Main(string[] args)
        {

            GetExcelData(1,1);
            

             static void GetExcelData(int i, int j)
            {
                string HeartRateSet1Path = "C:\\Users\\Alan Osefia\\Desktop\\DataGeneratorC#\\DataGeneratorCsharp\\DataGeneratorCsharp\\HeartRateSet1.xlsx";
                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(HeartRateSet1Path);
                Worksheet ws = wb.Worksheets[1];

     
                int numRows = 122-1;
                int numColumns = 4; 
                
                for (int x = 2; i <= numRows; x++)
                {
                    HeartRateData data = new HeartRateData();

   
                    var test = ws.Cells[x, 1].Value2;
                    data.SensorId = Convert.ToInt32(ws.Cells[x, 1].Value2);
                    data.HeartRateBPM = Convert.ToInt32(ws.Cells[x, 2].Value2);
                    data.TimeStamp =DateTime.Now;
                    data.Label = Convert.ToString(ws.Cells[x, 4].Value2);
                    data.Condition = Convert.ToString(ws.Cells[x, 5].Value2);

                    InsertHeartRateToDb(data);
                    Thread.Sleep(60000);
                   
                    i++;
                }

            }

            static void InsertHeartRateToDb(HeartRateData heartRate)
            {
                using (SqlConnection conn  = new SqlConnection())
                {
                    conn.ConnectionString = "Server=DESKTOP-QUHRVG6\\MSSQLServer02;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;";
                    conn.Open();
                    SqlCommand insertCommand = new SqlCommand("INSERT INTO dbo.HeartRateData (SensorId,HeartRateBPM,EnterTime) VALUES (@0, @1,@2)", conn);
                    insertCommand.Parameters.Add(new SqlParameter("0", heartRate.SensorId));
                    insertCommand.Parameters.Add(new SqlParameter("1", heartRate.HeartRateBPM));
                    insertCommand.Parameters.Add(new SqlParameter("2", heartRate.TimeStamp));
                    Console.WriteLine("Commands executed! Total rows affected are " + insertCommand.ExecuteNonQuery() + $"; sensorId={heartRate.SensorId} heartrate={heartRate.HeartRateBPM}, EnterTime={heartRate.TimeStamp}");
                
                }
            }
         
        }
    }
}