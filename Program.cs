using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Data.SQLite;
using ConsoleTableExt;

namespace EPPlusExcelToSQLDB
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Deleting Database if exists...");

            DeleteDatabase();

            Console.WriteLine("...Creating new Database...");

            CreateDatabase();

            ReadSpreadsheet();

            PrintDataBase();
        }




        private static void ReadSpreadsheet()
        {
            string dbFile = "URI=FILE:TempDatabase.db";

            Console.WriteLine("What is the file path to your Spreadsheet?");
            string file = Console.ReadLine();

            Console.WriteLine("Populating SQL Database and printing worksheets to console...");

            FileInfo fileInfo = new FileInfo(file);

            using (ExcelPackage package = new ExcelPackage(file))
            {
               

                foreach (var workSheet in package.Workbook.Worksheets)
                {

                    int totalRows = workSheet.Dimension.Rows;
                    int totalColumns = workSheet.Dimension.Columns;


                    List<Item> myList = new List<Item>();

                    for (int i = 2; i <= totalRows; i++)
                    {

                        myList.Add(new Item
                        {
                            Description = workSheet.Cells[i, 1].Value.ToString(),
                            Size = workSheet.Cells[i, 2].Value.ToString(),
                            Content = workSheet.Cells[i, 3].Value.ToString(),
                        });
                    }


                    for (int k = 0; k < myList.Count; k++)
                    {

                        SQLiteConnection connection = new SQLiteConnection(dbFile);
                        connection.Open();
                        string AddData = $"INSERT INTO tempdatabase (Description,Size,Content) VALUES ('{myList[k].Description}','{myList[k].Size}','{myList[k].Content}');";
                        SQLiteCommand command = new SQLiteCommand(AddData, connection);
                        command.ExecuteNonQuery();
                        connection.Close();

                    }

                    

                    

                    
                }
            }


        }
        static void CreateDatabase()
        {

            string dbFile = "URI=FILE:TempDatabase.db";
            SQLiteConnection connection = new SQLiteConnection(dbFile);
            connection.Open();

            string tbl = "CREATE TABLE tempdatabase (Description text, Size text, Content text);";
            SQLiteCommand command = new SQLiteCommand(tbl, connection);

            command.ExecuteNonQuery();
            connection.Close();
        }

        static void DeleteDatabase()
        {

            //add location of DB file here before running (after running first time)
            File.Delete(@"C:\Users\georg\Documents\CSharp\EPPlusExcelToSQLDB\TempDatabase.db");
        }

        static void PrintDataBase()
        {
            string dbFile = "URI=FILE:TempDatabase.db";
            var connection = new SQLiteConnection(dbFile);

            connection.Open();

            var tableCmd = connection.CreateCommand();
            tableCmd.CommandText = "SELECT * FROM tempdatabase";

            List<Item> tableData = new List<Item>();

            SQLiteDataReader reader = tableCmd.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    tableData.Add(
                    new Item
                    {
                        Description = new string(reader.GetString(0)),
                        Content = new string(reader.GetString(1)),
                        Size = new string(reader.GetString(2))
                    });
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            reader.Close();

            ConsoleTableBuilder
                .From(tableData)
                .WithTitle("DriveLogs", ConsoleColor.Green, ConsoleColor.Black)
                .WithTextAlignment(new Dictionary<int, TextAligntment>
                    {
                            {2, TextAligntment.Center}
                    })
                  .WithCharMapDefinition(new Dictionary<CharMapPositions, char> {
                        {CharMapPositions.BottomLeft, '=' },
                        {CharMapPositions.BottomCenter, '=' },
                        {CharMapPositions.BottomRight, '=' },
                        {CharMapPositions.BorderTop, '=' },
                        {CharMapPositions.BorderBottom, '=' },
                        {CharMapPositions.BorderLeft, '|' },
                        {CharMapPositions.BorderRight, '|' },
                        {CharMapPositions.DividerY, '|' },
                        {CharMapPositions.DividerX, '-' },
                    })

                .ExportAndWriteLine();

            Console.WriteLine("");
        }


    }
    public class Item
    {
        public string Description { get; set; }
        public string Size { get; set; }
        public string Content { get; set; }
    }

   
}

