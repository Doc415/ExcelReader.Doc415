using Dapper;
using Microsoft.Data.SqlClient;
using System.Dynamic;
using static OfficeOpenXml.ExcelErrorValue;

namespace ExcelReader.Doc415;

internal class DbHandler
{
    private string _connectionStringToDbUpdate = "Server=(localdb)\\MSSQLLocalDB; Integrated Security=true;";
    private string _connectionString = "Server=(localdb)\\MSSQLLocalDB;Initial Catalog=exceltodb; Integrated Security=true;";
    public DbHandler()
    {
        using (var connection = new SqlConnection(_connectionStringToDbUpdate))
        {
            connection.Open();
            try
            {
                string deleteDatabase = @"DROP DATABASE IF EXISTS exceltodb;";
                connection.Execute(deleteDatabase);
                Console.Write("db deleted");
                connection.Close();
            }
            catch
            {
                Console.WriteLine("no db to delete");
            }
        }

    }

    public void CreateDBTable(List<string> columnNames)
    {
        using (var connection = new SqlConnection(_connectionStringToDbUpdate))
        {
            connection.Open();
            string createDatabase = @"CREATE DATABASE exceltodb";
            connection.Execute(createDatabase);
            connection.Close();
        }
        using (var connection = new SqlConnection(_connectionString))
        {
            string createTable = "CREATE TABLE exceldata ([Id] INT IDENTITY(1,1) PRIMARY KEY (Id))";
            connection.Execute(createTable);

            foreach (var column in columnNames)
            {
                string addColumn = $"ALTER TABLE exceldata ADD {column} TEXT";
                connection.Execute(addColumn);
            }
        }
    }

    public void InsertRowsToDb(List<dynamic> rows)
    {
        var InsertData = MapRowsToQueryValues(rows);
        using (var connection = new SqlConnection(_connectionString))
        {
            string createTable = "";
           

            foreach (var column in InsertData.Item2)
            {
                string insertQuery = $"INSERT INTO exceldata ({InsertData.Item1}) VALUES ({column})";
                connection.Execute(insertQuery);
            }
        }
    }

    private (string,List<string>) MapRowsToQueryValues(List<dynamic> rows)
    {
        List<string> queryValues = new();
        string columnNamesForQuery = string.Join(",",ExcelFileHandler._colNames).TrimEnd(',');
        
        foreach (var  row in rows)
        {
            string values = "";
            var propertyBag = (IDictionary<string, object>)row;
            foreach (var property in propertyBag)
            {
                values += "'"+property.Value +"'"+ ",";

            }
            values=values.TrimEnd(',');
            queryValues.Add(values);
        }
        return (columnNamesForQuery, queryValues);
    }


}


