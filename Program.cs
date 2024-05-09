using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace CodeFirstDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Path to your Excel file
            string filePath = "D:\\Sohel vai\\Client.xlsx";
            List<Tuple<string, string, int, float>> dataList = new List<Tuple<string, string, int, float>>();
            // Check if the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File not found.");
                return;
            }
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Read data from the Excel file
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the first worksheet in the Excel file
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Get the number of rows and columns in the worksheet
                int rowCount = worksheet.Dimension.Rows;
                //int colCount = worksheet.Dimension.Columns;
               

                // Loop through each row in the worksheet
                for (int row = 2; row <= rowCount; row++)
                {
                   
                   
             
                        // Get the value of the cell at the current row and column
                        string Record_Id = worksheet.Cells[row, 1].Value.ToString();
                        string Client_Name = worksheet.Cells[row, 2].Value.ToString();

                     Object ob1 = worksheet.Cells[row, 3].Value;
                       
                    int Insight_Id = 0;
                    if(ob1!=null)
                    {
                        Insight_Id = int.Parse(ob1.ToString());
                    }
                    object ob = worksheet.Cells[row, 4].Value;
                    float Balance = 0.0f;
                    if (ob!=null)
                    {
                        if (float.TryParse(ob.ToString(), out float balanceValue))
                        {
                            // Conversion successful, assign the value
                            Balance = balanceValue;
                        }
                    }
                       // adding it to list tuple
                    dataList.Add(new Tuple<string, string, int,float>(Record_Id, Client_Name, Insight_Id, Balance));

                 
                }
            }


            Console.WriteLine(dataList.Count);
            // Retrieve connection string from app.config
            string connectionString = ConfigurationManager.ConnectionStrings["BlogDbContext"].ConnectionString;

            // Create SqlConnection object 
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    // Open the connection 
                   


                  //insert data into a table 
                    connection.Open();
                    string insertQuery = "INSERT INTO Client ([Record Id], [Client Name], [Report Insight Id], [Balance]) VALUES (@Value1, @Value2, @Value3, @Value4)";
                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                        command.Parameters.Add("@Value1", SqlDbType.NVarChar);
                        command.Parameters.Add("@Value2", SqlDbType.NVarChar);
                        command.Parameters.Add("@Value3", SqlDbType.Int);
                        command.Parameters.Add("@Value4", SqlDbType.Float);
                        for (int i = 0; i <dataList.Count; i++)
                            {
                            command.Parameters["@Value1"].Value = dataList[i].Item1;
                            command.Parameters["@Value2"].Value = dataList[i].Item2;
                            command.Parameters["@Value3"].Value = dataList[i].Item3;
                            command.Parameters["@Value4"].Value = dataList[i].Item4;
                            //Console.WriteLine(dataList[i].Item1 + " " + dataList[i].Item2 + " " + dataList[i].Item3 + " " + dataList[i].Item4);
                            // Execute the command for each item in the dataList
                            int rowsAffected = command.ExecuteNonQuery();
                            /*
                            if (rowsAffected > 0)
                            {
                                Console.WriteLine("Data is inserted successfully");
                            }
                            else
                            {
                                Console.WriteLine("Data failed to insert successfully");
                            }
                            */
                        }
                       
                    }
                       
                    connection.Close();






                    //read data from table
                    List<Tuple<string, string, int, float>> dataList1 = new List<Tuple<string, string, int, float>>();

                    connection.Open();

                    // Example: Execute a query
                    string query = "SELECT * FROM Client";
                    
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                // Process each row
                                // Example: Access data using reader.GetString(), reader.GetInt32(), etc.
                                // Access data from each column
                                string Id = reader.GetString(reader.GetOrdinal("Record Id"));
                                string name = reader.GetString(reader.GetOrdinal("Client Name"));
                                object ob1 = reader.GetOrdinal("Report Insight Id");
                                int Insight_Id = 0;
                                if (ob1 != null)
                                {
                                    Insight_Id = int.Parse(ob1.ToString());
                                }
                                object ob = reader.GetOrdinal("Balance");
                                float Balance = 0.0f;
                                if (ob != null)
                                {
                                    if (float.TryParse(ob.ToString(), out float balanceValue))
                                    {
                                        // Conversion successful, assign the value
                                        Balance = balanceValue;
                                    }
                                }
                                // Process the data as needed
                                dataList1.Add(new Tuple<string, string, int, float>(Id,name, Insight_Id, Balance));
                            }
                        }
                    }
                    Console.WriteLine(dataList1.Count);
                }
                catch (Exception ex)
                {
                    // Handle exceptions
                    Console.WriteLine("Error: " + ex.Message);
                }
                finally
                {
                    // Close the connection
                    connection.Close();
                    Console.WriteLine("Connection closed.");
                }
            }
        }

    }
}
