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
            DeleteDatabase();
            Console.WriteLine("DB Deleted");
            Console.ReadLine();
            CreateDatabase();
            ReadCsvFile();
            //ReadTable();
        }
        private static void ReadCsvFile()
        {
            string dbFile = "URI=FILE:TempDatabase.db";

            Console.WriteLine("What is the file path to your Csv?");
            string file = Console.ReadLine();
            FileInfo fileInfo = new FileInfo(file);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                foreach (var workSheet in package.Workbook.Worksheets)
                {
                    // Existing code

                    //ExcelWorksheet workSheet = package.Workbook.Worksheets["DriveLogSheet1"];


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


                        //Console.WriteLine(myList[k].Description + myList[k].Size + myList[k].Content);
                    }
                    ConsoleTableBuilder
                    .From(myList)
                    .WithTitle("DriveLog ", ConsoleColor.Green, ConsoleColor.Black)
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
            File.Delete(@"C:\Users\georg\Documents\CSharp\EPPlusExcelToSQLDB\TempDatabase.db");
        }

        
    }
    public class Item
    {
        public string Description { get; set; }
        public string Size { get; set; }
        public string Content { get; set; }
    }

    public class Item2
    {
        public string Description2 { get; set; }
        public string Size2 { get; set; }
        public string Content2 { get; set; }
    }
}

