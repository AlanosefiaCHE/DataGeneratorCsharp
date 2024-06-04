using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DatageneratorCsharp
{
    internal class Program
    {
        public class HeartRateData // use heartratedata class to make forwarding data to the database more consistent
        {
            public int? SensorId;
            public int? HeartRateBPM;
            public DateTime? TimeStamp;
            public bool IsProcessed;
            public bool IsFaulty;

            public HeartRateData(int sensorId, int heartRateBPM, DateTime? timeStamp)
            {
                SensorId = sensorId;
                HeartRateBPM = heartRateBPM;
                TimeStamp = timeStamp;
                IsProcessed = false;
                IsFaulty = false;

            }
        }
        static void Main(string[] args)
        {
            using SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Server=DESKTOP-NSA93TK\\SQLExpress;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;"; // laptop
            //conn.ConnectionString = "Server=DESKTOP-QUHRVG6\\MSSQLServer02;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;"; //computer
            conn.Open();

            PutHeartRateDataInDb(conn);

        }

        static void PutHeartRateDataInDb(SqlConnection conn)
        {
            const int startRow = 2;
            const int numRows = 5761;
            const int sensorIdIndex = 1;
            const int heartRateIndex = 2;
            const int oneMinuteInMs = 60000;

           string HeartRateSet1Path = "C:\\Users\\Gebruiker\\Desktop\\DatageneratorCsharp\\DataGeneratorCsharp\\HeartRateSet1.xlsx"; // This is excel path for my laptop.
           //string HeartRateSet1Path = "C:\\Users\\Alan Osefia\\Desktop\\DataGeneratorC#\\DataGeneratorCsharp\\DataGeneratorCsharp\\HeartRateSet1.xlsx"; // This is excel path for my computer.
            _Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(HeartRateSet1Path); // Workbook.Open opens the excel sheets
            Worksheet ws = wb.Worksheets[1];// Worksheets selects which worksheet you want to use
            Worksheet ws2 = wb.Worksheets[2];
            Worksheet ws3 = wb.Worksheets[3];

            for (int currentRow = startRow; currentRow < numRows; currentRow++) // For loop to put all the excel values into the class
            {
                HeartRateData heartRateDataRow = new HeartRateData(
                    sensorId: Convert.ToInt32(ws.Cells[currentRow, sensorIdIndex].Value2),
                    heartRateBPM: Convert.ToInt32(ws.Cells[currentRow, heartRateIndex].Value2),
                    timeStamp: DateTime.Now
                );
                //Second sensor
                HeartRateData heartRateDataRow2 = new HeartRateData(
                    sensorId: Convert.ToInt32(ws2.Cells[currentRow, sensorIdIndex].Value2),
                    heartRateBPM: Convert.ToInt32(ws2.Cells[currentRow, heartRateIndex].Value2),
                    timeStamp: DateTime.Now
                );
                
                HeartRateData heartRateDataRow3 = new HeartRateData(
                    sensorId: ConvertStringToIntOrNull(ws3.Cells[currentRow, sensorIdIndex].Value2),
                    heartRateBPM: ConvertStringToIntOrNull(ws3.Cells[currentRow, heartRateIndex].Value2),
                    timeStamp: RandomDateTimeOrNull()

                );
                InsertHeartRateToDb(conn,  heartRateDataRow);
                InsertHeartRateToDb(conn,  heartRateDataRow2);
                InsertHeartRateToDb(conn,  heartRateDataRow3);

                Thread.Sleep(10000);
            }
        }
        static int? ConvertStringToIntOrNull(dynamic? InputValue)
        {
            try
            {
                return Convert.ToInt32(InputValue);
            }
            catch (Exception e)
            {
                return null;
        
            }
        }

        static DateTime? RandomDateTimeOrNull()
        {
            Random random = new Random();
            int randomNumber = random.Next(4);
            return  randomNumber == 0 ? null : DateTime.Now;
        }

        static void InsertHeartRateToDb(SqlConnection conn, HeartRateData heartRate)
        {
            //Opens the connection with the Db
            SqlCommand insertCommand = new SqlCommand("INSERT INTO dbo.HeartRateData (HeartRateSensorId,HeartRateBPM,EnterTime,IsProcessed,IsFaulty) VALUES (@HeartRateSensorId, @HeartRateBPM,@TimeStamp, @IsProcessed,@IsFaulty)", conn);
            insertCommand.Parameters.Add(new SqlParameter("HeartRateSensorId", heartRate.SensorId));
            insertCommand.Parameters.Add(new SqlParameter("HeartRateBPM", heartRate.HeartRateBPM));
            insertCommand.Parameters.Add(new SqlParameter("TimeStamp", heartRate.TimeStamp.HasValue ? heartRate.TimeStamp.Value : DBNull.Value ));
            insertCommand.Parameters.Add(new SqlParameter("HeartRateBPM", heartRate.IsProcessed));
            insertCommand.Parameters.Add(new SqlParameter("HeartRateBPM", heartRate.IsFaulty));
            Console.WriteLine("Commands executed! Total rows affected are " + insertCommand.ExecuteNonQuery() + $"; sensorId={heartRate.SensorId} heartrate={heartRate.HeartRateBPM}, EnterTime={heartRate.TimeStamp}"); // Adds the data into the db, ands shows the enterd data in the console
        }
    }
}