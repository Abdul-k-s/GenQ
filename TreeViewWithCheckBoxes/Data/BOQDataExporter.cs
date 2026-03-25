using System;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

namespace GenQ.Data
{
    /// <summary>
    /// Exports BOQ data to SQL Server database
    /// </summary>
    public class BOQDataExporter
    {
        private readonly SqlServerConnection _connection;

        public BOQDataExporter(SqlServerConnection connection)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        }

        /// <summary>
        /// Initialize the database schema (creates tables if not exist)
        /// </summary>
        public void InitializeSchema()
        {
            const string createProjectsTable = @"
                IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='BOQ_Projects' AND xtype='U')
                CREATE TABLE BOQ_Projects (
                    ProjectId INT IDENTITY(1,1) PRIMARY KEY,
                    ProjectName NVARCHAR(255) NOT NULL,
                    ProjectNumber NVARCHAR(100),
                    RevitFilePath NVARCHAR(500),
                    ExportDate DATETIME2 DEFAULT GETDATE(),
                    ExportedBy NVARCHAR(100)
                )";

            const string createItemsTable = @"
                IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='BOQ_Items' AND xtype='U')
                CREATE TABLE BOQ_Items (
                    ItemId INT IDENTITY(1,1) PRIMARY KEY,
                    ProjectId INT NOT NULL,
                    ElementId INT NOT NULL,
                    Category NVARCHAR(100),
                    FamilyName NVARCHAR(255),
                    TypeName NVARCHAR(255),
                    Level NVARCHAR(100),
                    CSIDivision NVARCHAR(50),
                    CSIDescription NVARCHAR(255),
                    Quantity DECIMAL(18,4),
                    Unit NVARCHAR(50),
                    Area DECIMAL(18,4),
                    Volume DECIMAL(18,4),
                    Length DECIMAL(18,4),
                    Count INT DEFAULT 1,
                    Remarks NVARCHAR(500),
                    CONSTRAINT FK_BOQ_Items_Project FOREIGN KEY (ProjectId) 
                        REFERENCES BOQ_Projects(ProjectId) ON DELETE CASCADE
                )";

            const string createIndexes = @"
                IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_BOQ_Items_ProjectId')
                CREATE INDEX IX_BOQ_Items_ProjectId ON BOQ_Items(ProjectId);
                
                IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_BOQ_Items_Category')
                CREATE INDEX IX_BOQ_Items_Category ON BOQ_Items(Category);
                
                IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_BOQ_Items_CSIDivision')
                CREATE INDEX IX_BOQ_Items_CSIDivision ON BOQ_Items(CSIDivision);";

            _connection.ExecuteNonQuery(createProjectsTable);
            _connection.ExecuteNonQuery(createItemsTable);
            _connection.ExecuteNonQuery(createIndexes);
        }

        /// <summary>
        /// Create a new project record and return its ID
        /// </summary>
        public int CreateProject(string projectName, string projectNumber, string revitFilePath, string exportedBy)
        {
            const string sql = @"
                INSERT INTO BOQ_Projects (ProjectName, ProjectNumber, RevitFilePath, ExportedBy)
                OUTPUT INSERTED.ProjectId
                VALUES (@ProjectName, @ProjectNumber, @RevitFilePath, @ExportedBy)";

            var result = _connection.ExecuteScalar(sql,
                new SqlParameter("@ProjectName", projectName ?? (object)DBNull.Value),
                new SqlParameter("@ProjectNumber", projectNumber ?? (object)DBNull.Value),
                new SqlParameter("@RevitFilePath", revitFilePath ?? (object)DBNull.Value),
                new SqlParameter("@ExportedBy", exportedBy ?? Environment.UserName));

            return Convert.ToInt32(result);
        }

        /// <summary>
        /// Export BOQ items to database
        /// </summary>
        public void ExportItems(int projectId, DataTable boqData)
        {
            if (boqData == null || boqData.Rows.Count == 0)
                return;

            // Add ProjectId column if not present
            if (!boqData.Columns.Contains("ProjectId"))
            {
                boqData.Columns.Add("ProjectId", typeof(int));
                foreach (DataRow row in boqData.Rows)
                {
                    row["ProjectId"] = projectId;
                }
            }

            _connection.BulkInsert("BOQ_Items", boqData);
        }

        /// <summary>
        /// Export BOQ items asynchronously
        /// </summary>
        public async Task ExportItemsAsync(int projectId, DataTable boqData)
        {
            if (boqData == null || boqData.Rows.Count == 0)
                return;

            if (!boqData.Columns.Contains("ProjectId"))
            {
                boqData.Columns.Add("ProjectId", typeof(int));
                foreach (DataRow row in boqData.Rows)
                {
                    row["ProjectId"] = projectId;
                }
            }

            await _connection.BulkInsertAsync("BOQ_Items", boqData);
        }

        /// <summary>
        /// Insert a single BOQ item
        /// </summary>
        public void InsertItem(int projectId, BOQItem item)
        {
            const string sql = @"
                INSERT INTO BOQ_Items 
                (ProjectId, ElementId, Category, FamilyName, TypeName, Level, 
                 CSIDivision, CSIDescription, Quantity, Unit, Area, Volume, Length, Count, Remarks)
                VALUES 
                (@ProjectId, @ElementId, @Category, @FamilyName, @TypeName, @Level,
                 @CSIDivision, @CSIDescription, @Quantity, @Unit, @Area, @Volume, @Length, @Count, @Remarks)";

            _connection.ExecuteNonQuery(sql,
                new SqlParameter("@ProjectId", projectId),
                new SqlParameter("@ElementId", item.ElementId),
                new SqlParameter("@Category", item.Category ?? (object)DBNull.Value),
                new SqlParameter("@FamilyName", item.FamilyName ?? (object)DBNull.Value),
                new SqlParameter("@TypeName", item.TypeName ?? (object)DBNull.Value),
                new SqlParameter("@Level", item.Level ?? (object)DBNull.Value),
                new SqlParameter("@CSIDivision", item.CSIDivision ?? (object)DBNull.Value),
                new SqlParameter("@CSIDescription", item.CSIDescription ?? (object)DBNull.Value),
                new SqlParameter("@Quantity", item.Quantity),
                new SqlParameter("@Unit", item.Unit ?? (object)DBNull.Value),
                new SqlParameter("@Area", item.Area),
                new SqlParameter("@Volume", item.Volume),
                new SqlParameter("@Length", item.Length),
                new SqlParameter("@Count", item.Count),
                new SqlParameter("@Remarks", item.Remarks ?? (object)DBNull.Value));
        }

        /// <summary>
        /// Get all projects
        /// </summary>
        public DataTable GetProjects()
        {
            return _connection.ExecuteQuery("SELECT * FROM BOQ_Projects ORDER BY ExportDate DESC");
        }

        /// <summary>
        /// Get items for a specific project
        /// </summary>
        public DataTable GetProjectItems(int projectId)
        {
            return _connection.ExecuteQuery(
                "SELECT * FROM BOQ_Items WHERE ProjectId = @ProjectId ORDER BY Category, FamilyName",
                new SqlParameter("@ProjectId", projectId));
        }

        /// <summary>
        /// Delete a project and its items
        /// </summary>
        public void DeleteProject(int projectId)
        {
            _connection.ExecuteNonQuery(
                "DELETE FROM BOQ_Projects WHERE ProjectId = @ProjectId",
                new SqlParameter("@ProjectId", projectId));
        }
    }

    /// <summary>
    /// Represents a single BOQ item for database storage
    /// </summary>
    public class BOQItem
    {
        public int ElementId { get; set; }
        public string Category { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string Level { get; set; }
        public string CSIDivision { get; set; }
        public string CSIDescription { get; set; }
        public decimal Quantity { get; set; }
        public string Unit { get; set; }
        public decimal Area { get; set; }
        public decimal Volume { get; set; }
        public decimal Length { get; set; }
        public int Count { get; set; } = 1;
        public string Remarks { get; set; }
    }
}
