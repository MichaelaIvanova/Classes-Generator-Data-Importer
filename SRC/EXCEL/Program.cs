using LinqToExcel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace EXCEL
{
    class Program
    {

        static string path = @"C:\Users\michaela.ivanova\Desktop\EXCEL\EXCEL\Tests.xls";
        static string sheetName = "Test";
        static string pathJson = @"C:\Users\michaela.ivanova\Desktop\EXCEL\EXCEL\Test.json";
        static string directory = @"C:\Users\michaela.ivanova\Desktop\EXCEL\EXCEL\";
        static List<Test> results = new List<Test>();
        static DateTime lastRead = DateTime.MinValue;
        static FileSystemWatcher watcher = new FileSystemWatcher();

        static void Main(string[] args)
        {
            string pathToAssembly = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var rootDirectory = System.IO.Path.GetDirectoryName(path);

            ImportData(path);
            //Console.WriteLine("Watching started");
            //CreateFileWatcher(directory);

            //Console.Read();
        }

        public static void ImportData(string path)
        {
            string con =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" +
            @"Extended Properties='Excel 12.0 Xml;HDR=YES;'";

            var list = new List<Person>();

            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Test$]", connection);

                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string name = dr[0].ToString();
                        string lastName = dr[1].ToString();
                        string age = dr[2].ToString();

                        var personToadd = new Person { Fname = name, Lname = lastName, Age = age };
                        list.Add(personToadd);
                        Console.WriteLine(name + " " + lastName + " "+age);
                    }

                }
            }

            var jsonCollection = Newtonsoft.Json.JsonConvert.SerializeObject(list);
            Console.WriteLine(jsonCollection);
        }

        public static void ImportDataExcelLinq(string pathToEcxelFile, string sheetName, string pathToOutputJsonFile)
        {
            var excelFile = new ExcelQueryFactory(pathToEcxelFile);
            var data = excelFile.Worksheet(sheetName).Select(i => i).ToList();
            var colums = excelFile.GetColumnNames(sheetName).ToList();

            var builder = new StringBuilder();

            builder.Append("[");

            foreach (var item in data)
            {
                var cellsPerRow = item.Count;
                builder.Append("{");
                for (var i = 0; i < cellsPerRow; i++)
                {

                    builder.AppendFormat("\"{0}\":\"{1}\"", String.Join("", colums[i].ToString().Split(new char[] { ' ' })), item[i].ToString());
                    builder.Append(",");
                }
                builder.Length--;
                builder.Append("}");
                builder.Append(",");
            }
            //remove the last ,
            builder.Length--;
            builder.Append("]");

            using (StreamWriter w = new StreamWriter(pathToOutputJsonFile))
            {
                string json = builder.ToString();
                w.Write(json);
            }

        }

        public static void CreateFileWatcher(string path)
        {
            
            watcher.Path = path;

            watcher.NotifyFilter = NotifyFilters.LastWrite;

            //watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.LastWrite
            //   | NotifyFilters.FileName | NotifyFilters.DirectoryName;

            watcher.Filter = "Tests.xls";

            // Add event handlers.
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.EnableRaisingEvents = true;
            
        }

        // Define the event handlers.
        private static void OnChanged(object source, FileSystemEventArgs e)
        {
            Console.WriteLine("OnChange event fired" + e.ChangeType);
            ImportDataExcelLinq(path, sheetName, pathJson);
            Console.WriteLine("Data filled...");
            Console.WriteLine("Parsing json .....");
            results.Clear();
            results = LoadJson(pathJson);

            if (results.Count != 0)
            {
                foreach (var r in results)
                {
                    Console.WriteLine("{0} > {1} > {2}", r.CustomPropertyFirstName, r.CustomPropertyLastName, r.CustomPropertyAge);
                }
            }

        }


        private static List<Test> LoadJson(string path)
        {
            using (StreamReader r = new StreamReader(path))
            {
                string json = r.ReadToEnd();
                List<Test> items = JsonConvert.DeserializeObject<List<Test>>(json);

                try
                {
                    items = items.Distinct().ToList();
                }
                catch (Exception ex)
                {

                }

                return items;
            }
        }
    }
}
