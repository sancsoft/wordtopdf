using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPDF.Service
{
    public interface IDocumentQueue
    {
        string SourceName();
        int Count();
        DocumentTarget NextDocument();
        void CompleteDocument(DocumentTarget documentTarget);
    }
}
