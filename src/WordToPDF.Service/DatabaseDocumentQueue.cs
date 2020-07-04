using System.Data;
using Dapper;

namespace WordToPDF.Service
{
    public class DatabaseDocumentQueue : IDocumentQueue
    {
        protected IDbConnection _connection;
        protected string _tableName;

        public DatabaseDocumentQueue(IDbConnection connection, string tableName)
        {
            _connection = connection;
            _tableName = tableName;
        }

        public string SourceName()
        {
            return _tableName;
        }

        public int Count()
        {
            return _connection.ExecuteScalar<int>($"SELECT COUNT(Id) FROM {_tableName} WHERE ResultCode = -1");
        }

        public DocumentTarget NextDocument()
        {
            return _connection.QueryFirst<DocumentTarget>($"SELECT * FROM {_tableName} WHERE ResultCode = -1 ORDER BY Id");
        }

        public void CompleteDocument(DocumentTarget documentTarget)
        {
            _connection.Execute($"UPDATE {_tableName} SET ResultCode=@ResultCode, InputFile=@InputFile, OutputFile=@OutputFile", documentTarget);
        }
    }
}
