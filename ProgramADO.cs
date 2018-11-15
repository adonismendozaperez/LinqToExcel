using System;
using System.Linq;
using LinqToExcel.Attributes;
using System.Data.SqlClient;

namespace LinqToExcel
{
    class ProgramADO
    {
        static void Main(string[] args)
        {
            //Set route of file
            var excelFile = new ExcelQueryFactory(@"C:\Users\Dell\Desktop\test.xlsx");

            //Query for get all data to the file xsl
            var Persons = from p in excelFile.Worksheet<Person>()
                          select p;

            //Set Connection String of your Database
            SqlConnection Connection = new SqlConnection(@"server=SQLEXPRESS; database=db-LinqToExcel; integrated security = true");
            Connection.Open();

            Console.WriteLine("Processing...");
            foreach (Person person in Persons)
            {
                String Script = @"INSERT INTO Person(Name, Age) values ('" + person.Name + "'," + person.Age + ")";
                SqlCommand Command = new SqlCommand(Script, Connection);
                Command.ExecuteNonQuery();
            }
            Connection.Close();
            Console.WriteLine("Ended process.");
        }

        public class Person
        {
            [ExcelColumn("Name")]
            public string Name { get; set; }

            [ExcelColumn("Age")]
            public int Age { get; set; }
        }
    }
}