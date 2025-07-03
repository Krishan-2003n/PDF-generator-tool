using Dapper;
using PDF_generator_tool.Models;
using System.Data.SqlClient;

namespace PDF_generator_tool.Services
{
    public class ExportService : IExportService
    {
        public async Task<List<string>> GetTablesAsync(string connectionString)
        {
            using var conn = new SqlConnection(connectionString);
            var tables = await conn.QueryAsync<string>("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'");
            return tables.ToList();
        }

        public async Task<List<string>> GetColumnsAsync(string connectionString, string tableName)
        {
            using var conn = new SqlConnection(connectionString);
            var columns = await conn.QueryAsync<string>(
                "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @Table",
                new { Table = tableName });
            return columns.ToList();
        }

        public async Task<List<Dictionary<string, object>>> FetchDataAsync(string connectionString, string tableName, List<string> fields)
        {
            using var conn = new SqlConnection(connectionString);
            var fieldList = string.Join(",", fields.Select(f => $"[{f}]"));
            var query = $"SELECT {fieldList} FROM [{tableName}]";
            var result = await conn.QueryAsync(query);

            return result.Select(r => (IDictionary<string, object>)r)
                         .Select(d => d.ToDictionary(k => k.Key, v => v.Value))
                         .ToList();
        }

        public async Task<List<string>> GetPrimaryKeys(string connectionString, string tableName)
        {
            using var connection = new SqlConnection(connectionString);
            var query = @"
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS TC
        INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU
            ON TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME
        WHERE TC.CONSTRAINT_TYPE = 'PRIMARY KEY'
        AND KU.TABLE_NAME = @TableName";

            var primaryKeys = (await connection.QueryAsync<string>(query, new { TableName = tableName })).ToList();
            return primaryKeys;
        }

        public async Task<List<string>> GetPrimaryKeysAsync(string connectionString, string tableName)
        {
            using var connection = new SqlConnection(connectionString);
            var query = @"SELECT COLUMN_NAME
                  FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
                  WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + QUOTENAME(CONSTRAINT_NAME)), 'IsPrimaryKey') = 1
                  AND TABLE_NAME = @tableName";

            var keys = await connection.QueryAsync<string>(query, new { tableName });
            return keys.ToList();
        }

        public async Task<List<Dictionary<string, object>>> FetchPreviewDataAsync(string connectionString, string tableName, List<string> fields, int topRows = 5)
        {
            using var conn = new SqlConnection(connectionString);
            var fieldList = string.Join(",", fields.Select(f => $"[{f}]"));
            var query = $"SELECT TOP {topRows} {fieldList} FROM [{tableName}]";
            var result = await conn.QueryAsync(query);

            return result.Select(r => (IDictionary<string, object>)r)
                         .Select(d => d.ToDictionary(k => k.Key, v => v.Value))
                         .ToList();
        }


    }
}
