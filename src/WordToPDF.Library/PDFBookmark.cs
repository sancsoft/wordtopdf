using System.Collections.Generic;

namespace WordToPDF.Library
{
    class PDFBookmark
    {
        public int page { get; set; }
        public string title { get; set; }
        public List<PDFBookmark> children { get; set; }
    }
}
