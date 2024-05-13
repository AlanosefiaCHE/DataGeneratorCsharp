using System;
using System.Data.SqlClient;

namespace DatageneratorCsharp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (SqlConnection conn = new SqlConnection())
            {
                conn.ConnectionString = "Server=DESKTOP-QUHRVG6\\MSSQLServer02;Database=APPS;Trusted_Connection=True;TrustServerCertificate=true;";
                conn.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM dbo.Medication", conn);
                // using the code here...
               using (SqlDataReader reader = command.ExecuteReader()) 
                { 
                    while (reader.Read()) 
                    {
                        Console.WriteLine(String.Format("{0} \t | {1} \t | {2}",
                         // call the objects from their index
                         reader[0], reader[1], reader[2]));
                    }
                }
               

            }
            Console.WriteLine("Hello World!");
        }
    }
}