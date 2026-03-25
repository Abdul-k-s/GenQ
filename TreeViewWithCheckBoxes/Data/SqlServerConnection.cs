using System;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

namespace GenQ.Data
{
    /// <summary>
    /// SQL Server connection helper for GenQ BOQ output
    /// </summary>
    public class SqlServerConnection : IDisposable
    {
        private readonly string _connectionString;
        private SqlConnection _connection;
        private bool _disposed;

        /// <summary>
        /// Initialize with connection string
        /// </summary>
        /// <param name="connectionString">SQL Server connection string</param>
        public SqlServerConnection(string connectionString)
        {
            _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        }

        /// <summary>
        /// Initialize with individual parameters
        /// </summary>
        public SqlServerConnection(string server, string database, string username = null, string password = null, bool integratedSecurity = true)
        {
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = server,
                InitialCatalog = database,
                IntegratedSecurity = integratedSecurity,
                TrustServerCertificate = true,
                ConnectTimeout = 30
            };

            if (!integratedSecurity && !string.IsNullOrEmpty(username))
            {
                builder.UserID = username;
                builder.Password = password;
            }

            _connectionString = builder.ConnectionString;
        }

        /// <summary>
        /// Connection string being used
        /// </summary>
        public string ConnectionString => _connectionString;

        /// <summary>
        /// Open the database connection
        /// </summary>
        public void Open()
        {
            if (_connection == null)
            {
                _connection = new SqlConnection(_connectionString);
            }

            if (_connection.State != ConnectionState.Open)
            {
                _connection.Open();
            }
        }

        /// <summary>
        /// Open the database connection asynchronously
        /// </summary>
        public async Task OpenAsync()
        {
            if (_connection == null)
            {
                _connection = new SqlConnection(_connectionString);
            }

            if (_connection.State != ConnectionState.Open)
            {
                await _connection.OpenAsync();
            }
        }

        /// <summary>
        /// Close the database connection
        /// </summary>
        public void Close()
        {
            if (_connection != null && _connection.State != ConnectionState.Closed)
            {
                _connection.Close();
            }
        }

        /// <summary>
        /// Test the connection
        /// </summary>
        public bool TestConnection(out string errorMessage)
        {
            errorMessage = null;
            try
            {
                using (var testConn = new SqlConnection(_connectionString))
                {
                    testConn.Open();
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Execute a non-query command (INSERT, UPDATE, DELETE)
        /// </summary>
        public int ExecuteNonQuery(string sql, params SqlParameter[] parameters)
        {
            EnsureOpen();
            using (var cmd = new SqlCommand(sql, _connection))
            {
                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Execute a non-query command asynchronously
        /// </summary>
        public async Task<int> ExecuteNonQueryAsync(string sql, params SqlParameter[] parameters)
        {
            await EnsureOpenAsync();
            using (var cmd = new SqlCommand(sql, _connection))
            {
                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }
                return await cmd.ExecuteNonQueryAsync();
            }
        }

        /// <summary>
        /// Execute a scalar query (returns single value)
        /// </summary>
        public object ExecuteScalar(string sql, params SqlParameter[] parameters)
        {
            EnsureOpen();
            using (var cmd = new SqlCommand(sql, _connection))
            {
                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }
                return cmd.ExecuteScalar();
            }
        }

        /// <summary>
        /// Execute a query and return a DataTable
        /// </summary>
        public DataTable ExecuteQuery(string sql, params SqlParameter[] parameters)
        {
            EnsureOpen();
            using (var cmd = new SqlCommand(sql, _connection))
            {
                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }

                using (var adapter = new SqlDataAdapter(cmd))
                {
                    var table = new DataTable();
                    adapter.Fill(table);
                    return table;
                }
            }
        }

        /// <summary>
        /// Execute a query asynchronously and return a DataTable
        /// </summary>
        public async Task<DataTable> ExecuteQueryAsync(string sql, params SqlParameter[] parameters)
        {
            await EnsureOpenAsync();
            using (var cmd = new SqlCommand(sql, _connection))
            {
                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }

                var table = new DataTable();
                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    table.Load(reader);
                }
                return table;
            }
        }

        /// <summary>
        /// Begin a transaction
        /// </summary>
        public SqlTransaction BeginTransaction()
        {
            EnsureOpen();
            return _connection.BeginTransaction();
        }

        /// <summary>
        /// Bulk insert data from a DataTable
        /// </summary>
        public void BulkInsert(string tableName, DataTable data, SqlTransaction transaction = null)
        {
            EnsureOpen();
            using (var bulkCopy = new SqlBulkCopy(_connection, SqlBulkCopyOptions.Default, transaction))
            {
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.BatchSize = 1000;
                bulkCopy.BulkCopyTimeout = 300;

                // Map columns
                foreach (DataColumn column in data.Columns)
                {
                    bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                }

                bulkCopy.WriteToServer(data);
            }
        }

        /// <summary>
        /// Bulk insert data asynchronously
        /// </summary>
        public async Task BulkInsertAsync(string tableName, DataTable data, SqlTransaction transaction = null)
        {
            await EnsureOpenAsync();
            using (var bulkCopy = new SqlBulkCopy(_connection, SqlBulkCopyOptions.Default, transaction))
            {
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.BatchSize = 1000;
                bulkCopy.BulkCopyTimeout = 300;

                foreach (DataColumn column in data.Columns)
                {
                    bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                }

                await bulkCopy.WriteToServerAsync(data);
            }
        }

        private void EnsureOpen()
        {
            if (_connection == null || _connection.State != ConnectionState.Open)
            {
                Open();
            }
        }

        private async Task EnsureOpenAsync()
        {
            if (_connection == null || _connection.State != ConnectionState.Open)
            {
                await OpenAsync();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    Close();
                    _connection?.Dispose();
                }
                _disposed = true;
            }
        }
    }
}
