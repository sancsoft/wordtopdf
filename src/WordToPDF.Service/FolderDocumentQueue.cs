using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPDF.Service
{
    public class FolderDocumentQueue: IDocumentQueue
    {
        protected string _sourceName;
        protected string _watchPath;
        protected int _documentIndex;

        public FolderDocumentQueue(string watchPath, string sourceName = "")
        {
            _watchPath = watchPath;
            _sourceName = (string.IsNullOrEmpty(_sourceName)) ? "Folder:" + watchPath : sourceName;
            _documentIndex = 0;
        }

        public string SourceName()
        {
            return _sourceName;
        }

        public int Count()
        {
            int count = 0;
            try
            {
                string[] files = Directory.GetFiles(_watchPath, "*.docx");
                count = files.Count();
            }
            catch (Exception e)
            {
                Log.Warning("Unable to read source directory {0} - {1}", _watchPath, e.Message);
            }
            return count;
        }

        public DocumentTarget NextDocument()
        {
            DocumentTarget documentTarget = null;
            try
            {
                string[] files = Directory.GetFiles(_watchPath, "*.docx");
                if (files.Count() > 0)
                {
                    documentTarget = new DocumentTarget()
                    { 
                        Id = _documentIndex++, 
                        InputFile = files[0],
                        OutputFile = "",
                        ResultCode = 0
                    };
                }
            }
            catch (Exception e)
            {
                Log.Warning("Unable to read source directory {0} - {1}", _watchPath, e.Message);
            }
            return documentTarget;
        }

        public void CompleteDocument(DocumentTarget documentTarget)
        {
            try
            {
                File.Delete(documentTarget.InputFile);
            }
            catch (Exception e)
            {
                Log.Warning("Unable to complete document {0}:{1} - {2}", documentTarget.Id, documentTarget.InputFile, e.Message);
            }
        }
    }
}
