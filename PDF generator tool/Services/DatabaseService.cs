using System.Data.SqlClient;
using Dapper;

namespace PDF_generator_tool.Services
{
    public class DatabaseService
    {
        private string ConnectionString;

        public DatabaseService(string connectionString)
        {
            ConnectionString = connectionString;
        }

        public List<string> GetAllTables()
        {
            using var connection = new SqlConnection(ConnectionString);
            var query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'";
            return connection.Query<string>(query).ToList();
        }

        public List<string> GetColumnsByTable(string tableName)
        {
            using var connection = new SqlConnection(ConnectionString);
            var query = @"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TableName";
            return connection.Query<string>(query, new { TableName = tableName }).ToList();
        }

        public List<Dictionary<string, object>> GetTableData(string tableName)
        {
            var result = new List<Dictionary<string, object>>();
            using var connection = new SqlConnection(ConnectionString);
            var cmd = connection.CreateCommand();
            cmd.CommandText = $"SELECT * FROM [{tableName}]";
            connection.Open();
            var reader = cmd.ExecuteReader();

            Console.WriteLine("Fetching data from table: " + tableName);

            while (reader.Read())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var value = reader.GetValue(i);
                    Console.WriteLine($"{columnName}: {value}");
                    row[columnName] = value;
                }
                result.Add(row);
            }

            return result;
        }

        public async Task<List<string>> GetPrimaryKeys(string connectionString, string tableName)
        {
            using var conn = new SqlConnection(connectionString);
            var query = @"
        SELECT kcu.COLUMN_NAME
        FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
        JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu
            ON tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME
        WHERE tc.TABLE_NAME = @TableName AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'";

            var keys = await conn.QueryAsync<string>(query, new { TableName = tableName });
            return keys.ToList();
        }

    }
}
