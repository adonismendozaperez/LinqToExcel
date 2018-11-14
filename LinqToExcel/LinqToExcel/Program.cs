using LinqToExcel.Attributes;
using System;
using System.Data.SqlClient;
using System.Linq;

namespace LinqToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Set route of file
            var excel = new ExcelQueryFactory(@"C:\Users\Dell\Desktop\test.xlsx");

            //Query for get all data to the file xsl
            var Persons = from p in excel.Worksheet<Person>()
                          select p;

            //Set Connection String of your Database
            SqlConnection conexion = new SqlConnection(@"server=SQLEXPRESS; database=db-LinqToExcel; integrated security = true");
            conexion.Open();
            string cadena = "";

            Console.WriteLine("Procesando...");
            foreach (Person person in Persons)
            {
                cadena = @"INSERT INTO Person(Name, Age) values ('" + person.Name + "'," + person.Age + ")";
                SqlCommand comando = new SqlCommand(cadena, conexion);
                comando.ExecuteNonQuery();
            }
            conexion.Close();
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