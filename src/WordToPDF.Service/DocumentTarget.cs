namespace WordToPDF.Service
{
    public class DocumentTarget
    {
        public int Id { get; set; }
        public string InputFile { get; set; }
        public string OutputFile { get; set; }
        public int ResultCode { get; set; }
    }
}
