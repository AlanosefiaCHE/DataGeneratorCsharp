using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DatageneratorCsharp
{
    internal class Program
    {
        public class HeartRateData // use heartratedata class to make forwarding data to the database more consistent
        {
            public int SensorId;
            public int HeartRateBPM;
            public DateTime TimeStamp;

            public HeartRateData(int sensorId, int heartRateBPM)
            {
                SensorId = sensorId;
                HeartRateBPM = heartRateBPM;
                TimeStamp = DateTime.Now;
            }
        }
        static void Main(string[] args)
        {
            using SqlConnection conn = new SqlConnection();
            //conn.ConnectionString = "Server=DESKTOP-NSA93TK\\SQLExpress;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;";
            conn.ConnectionString = "Server=DESKTOP-QUHRVG6\\MSSQLServer02;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;";
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

           // string HeartRateSet1Path = "C:\\Users\\Gebruiker\\Desktop\\DatageneratorCsharp\\DataGeneratorCsharp\\HeartRateSet1.xlsx"; // This is excel path for my PC.
            string HeartRateSet1Path = "C:\\Users\\Alan Osefia\\Desktop\\DataGeneratorC#\\DataGeneratorCsharp\\DataGeneratorCsharp\\HeartRateSet1.xlsx"; // This is excel path for my PC.
            _Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(HeartRateSet1Path); // Workbook.Open opens the excel sheets
            Worksheet ws = wb.Worksheets[1];// Worksheets selects which worksheet you want to use
            Worksheet ws2 = wb.Worksheets[2];

            for (int currentRow = startRow; currentRow < numRows; currentRow++) // For loop to put all the excel values into the class
            {
                HeartRateData heartRateDataRow = new HeartRateData(
                    sensorId: Convert.ToInt32(ws.Cells[currentRow, sensorIdIndex].Value2),
                    heartRateBPM: Convert.ToInt32(ws.Cells[currentRow, heartRateIndex].Value2)
                );
                HeartRateData heartRateDataRow2 = new HeartRateData(
                    sensorId: Convert.ToInt32(ws2.Cells[currentRow, sensorIdIndex].Value2),
                    heartRateBPM: Convert.ToInt32(ws2.Cells[currentRow, heartRateIndex].Value2)
                );
                InsertHeartRateToDb(conn,  heartRateDataRow);
                InsertHeartRateToDb(conn,  heartRateDataRow2);

                Thread.Sleep(oneMinuteInMs);
            }
        }
        static void InsertHeartRateToDb(SqlConnection conn, HeartRateData heartRate)
        {
            //Opens the connection with the Db
            SqlCommand insertCommand = new SqlCommand("INSERT INTO dbo.HeartRateData (SensorId,HeartRateBPM,EnterTime) VALUES (@SensorId, @HeartRateBPM,@TimeStamp)", conn);
            insertCommand.Parameters.Add(new SqlParameter("SensorId", heartRate.SensorId));
            insertCommand.Parameters.Add(new SqlParameter("HeartRateBPM", heartRate.HeartRateBPM));
            insertCommand.Parameters.Add(new SqlParameter("TimeStamp", heartRate.TimeStamp));
            Console.WriteLine("Commands executed! Total rows affected are " + insertCommand.ExecuteNonQuery() + $"; sensorId={heartRate.SensorId} heartrate={heartRate.HeartRateBPM}, EnterTime={heartRate.TimeStamp}"); // Adds the data into the db, ands shows the enterd data in the console
        }
    }
}