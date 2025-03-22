using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HttpListenerService.Model
{
   public class WordDocumentInputModel
    { 
        public string FileToDownload {  get; set; }
        public string MyDocumentsFolderPath {  get; set; }
        public string AddInsConfigFileName {  get; set; }
        public string TempFileName {  get; set; }
    }
}
