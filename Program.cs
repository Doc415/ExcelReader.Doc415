namespace ExcelReader.Doc415
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DbHandler _db = new();
            ExcelFileHandler _fileHandler = new ExcelFileHandler();
            _db.CreateDBTable(_fileHandler.GetColumnNames());
            var rows = _fileHandler.GetRows();
            _db.InsertRowsToDb(rows);        }
    }
}
