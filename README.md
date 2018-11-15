# LinqToExcel
This is an example using the [LinqToExcel](https://github.com/paulyoder/LinqToExcel) library, to read an Excel file and insert it into an MSQL database.

## Prerequisite to use
You must install [Microsoft Access Database Engine 2010 Redistributable](https://www.microsoft.com/en-in/download/details.aspx?id=13255).

Now next step is Install NuGet Package:
> Install-Package LinqToExcel -Version 1.11.0

If you use .NET CLI Install:
> dotnet add package LinqToExcel --version 1.11.0

## Use
In this step I get the excel file path, and I make a query and get the excel file records.

```cSharp
//Set route of file
var excelFile = new ExcelQueryFactory(@"C:\Users\Dell\Desktop\test.xlsx");

//Query for get all data to the file xsl
var Persons = from p in excelFile.Worksheet<Person>()
                select p;
```
In this step I will use **ADO .NET** I create the database connection, I open the connection, and I go through my query named Persons, and I insert each of these to my database.

```cSharp
//Set Connection String of your Database
SqlConnection Connection = new SqlConnection(@"server=SQLEXPRESS; database=db-LinqToExcel; integrated security = true");
Connection.Open();

foreach (Person person in Persons)
{
    String Script = @"INSERT INTO Person(Name, Age) values ('" + person.Name + "'," + person.Age + ")";
    SqlCommand Command = new SqlCommand(Script, Connection);
    Command.ExecuteNonQuery();
}
Connection.Close();
```
## Complete example
[Exemple in ADO.NET](https://github.com/adonismendozaperez/LinqToExcel/blob/master/ProgramADO.cs)

## Reference
[LinqToExcel](https://github.com/paulyoder/LinqToExcel)