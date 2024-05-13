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
                    Console.WriteLine("Id\tName\t\tAmountMg");
                    while (reader.Read()) 
                    {
                        Console.WriteLine(String.Format("{0} \t | {1} \t | {2}",
                         // call the objects from their index
                         reader[0], reader[1], reader[2]));
                    }
                }
                Console.WriteLine("Data displayed! Now press enter to move to the next section!");
                Console.ReadLine();
                Console.Clear();

                Console.WriteLine("INSERT INTO command");

                SqlCommand insertCommand = new SqlCommand("INSERT INTO dbo.Medication (Name, AmountMg) VALUES (@0, @1)", conn);
                insertCommand.Parameters.Add(new SqlParameter("0", "MDMA"));
                insertCommand.Parameters.Add(new SqlParameter("1", "200"));

                Console.WriteLine("Commands executed! Total rows affected are " + insertCommand.ExecuteNonQuery());
                Console.WriteLine("Done! Press enter to move to the next step");
                Console.ReadLine();
                Console.Clear();

            }
            
        }
    }
}