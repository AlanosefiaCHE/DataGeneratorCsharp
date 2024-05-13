using System;
using System.Data.SqlClient;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
//https://github.com/ExcelDataReader/ExcelDataReader
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace DatageneratorCsharp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        public class HeartRateData
        {
            public int sensorId;
            public int HeartRateBPM;
            public DateTime TimeStamp;
        }

        static void Main(string[] args)
        {

            GetExcelData(0,0);
            

             static string GetExcelData(int i, int j)
            {
                string HeartRateSet1Path = "C:\\Users\\Alan Osefia\\Desktop\\DataGeneratorC#\\DataGeneratorCsharp\\DataGeneratorCsharp\\HeartRateSet1.xlsx";
                _Application excel = new _Excel.Application();
                Workbook wb = excel.Workbooks.Open(HeartRateSet1Path);
                Worksheet ws = wb.Worksheets[1];

                i++;
                j++;

                if (ws.Cells[i,j].Value2 != null)
                {
                    Console.WriteLine(ws.Cells[i, j].Value2);
                }
                return "err";
            }

            //static void InsertIntoDb()

            //{
            //    using (SqlConnection conn = new SqlConnection())
            //    {
            //        conn.ConnectionString = "Server=DESKTOP-QUHRVG6\\MSSQLServer02;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;";
            //        conn.Open();

            //        string heartRateDataPath = "HeartRateSet1.xlsx";
            //        Timer timer = new Timer(TimerCallback, null, TimeSpan.Zero, TimeSpan.FromMinutes(1));

            //        static void TimerCallback(object state)
            //        {
            //            string heartRateDataPath = "HeartRateSet1.xlsx";
            //            string[] rowHeartRateData = ReadRowFromExcel(heartRateDataPath);

            //            AddHeartRateToDb(rowHeartRateData);

            //        }
            //        static string[] ReadRowFromExcel(string excelPath)
            //        {
            //            string[] empty = new string[0];
            //            using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            //            {
            //                using (var reader = ExcelReaderFactory.CreateReader(stream))
            //                {
            //                    var dataSet = reader.AsR
    
            //            }
            //            }
            //            return empty;
            //        }

            //        static void AddHeartRateToDb(string[] heartRateData)
            //        {

            //        }



            //        SqlCommand command = new SqlCommand("SELECT * FROM dbo.Medication", conn);
            //        // using the code here...
            //        using (SqlDataReader reader = command.ExecuteReader())
            //        {
            //            Console.WriteLine("Id\tName\t\tAmountMg");
            //            while (reader.Read())
            //            {
            //                Console.WriteLine(String.Format("{0} \t | {1} \t | {2}",
            //                 // call the objects from their index
            //                 reader[0], reader[1], reader[2]));
            //            }
            //        }
            //        Console.WriteLine("Data displayed! Now press enter to move to the next section!");
            //        Console.ReadLine();
            //        Console.Clear();

            //        Console.WriteLine("INSERT INTO command");

            //        SqlCommand insertCommand = new SqlCommand("INSERT INTO dbo.Medication (Name, AmountMg) VALUES (@0, @1)", conn);
            //        insertCommand.Parameters.Add(new SqlParameter("0", "MDMA"));
            //        insertCommand.Parameters.Add(new SqlParameter("1", "200"));

            //        Console.WriteLine("Commands executed! Total rows affected are " + insertCommand.ExecuteNonQuery());
            //        Console.WriteLine("Done! Press enter to move to the next step");
            //        Console.ReadLine();
            //        Console.Clear();

            //    }
            //}
           
            
        }
    }
}