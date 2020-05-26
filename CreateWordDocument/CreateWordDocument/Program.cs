using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            var dataSets = TestData.CreateDocumentDataSets();

            var generator = new DocumentGenerator();
            generator.DocumentDataSets = dataSets;
            generator.DocumentName = "DocumentSample.docx";
            generator.Generate();
        }
    }
}
