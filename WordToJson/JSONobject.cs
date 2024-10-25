using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToJson
{
    class JSONobject
    {
        public string document_title { get; set; }
        public string ?section_title { get; set; }
        public string[] ?content {  get; set; }
        public string[] ?tables { get; set; }

        public JSONobject(string docTitle,string ?sectionTitle)
        {
            this.document_title = docTitle;
            this.section_title = sectionTitle;
        }
    }
}
