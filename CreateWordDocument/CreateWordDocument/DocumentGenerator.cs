using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class DocumentGenerator
    {
        public string OutputFolder { get; set; }
        public string DocumentName { get; set; }
        public IList<DocumentDataSet> DocumentDataSets { get; set; }

        public string DocumentFullName
        {
            get
            {
                return OutputFolder + DocumentName;
            }
        }

        public DocumentGenerator()
        {
            OutputFolder = GlobalConstants.GenerateFolderPath;
        }

        public void Generate()
        {
            GlobalHelper.CreateFolderIfNotExists(OutputFolder);

            FileInfo newFile = new FileInfo(DocumentFullName);
            
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(DocumentFullName);
            }

            DocumentSetting setting = new DocumentSetting()
            {
                PaperType = PaperType.A4_V,
                PaperMarType = PaperMarType.MarType1,
                SavePath = DocumentFullName,
                Footer = DocumentName
            };

            var wordTables = new List<WordTable>();
            foreach (var ws in DocumentDataSets)
            {
                if (ws.WorkSheetType == WorkSheetTypeEnum.Summary)
                {
                    var wordTable = WordTableGenerator.GetSummaryWordTable(ws);
                    wordTables.Add(wordTable);
                }
                else if (ws.WorkSheetType == WorkSheetTypeEnum.ITD)
                {
                    var wordTable = WordTableGenerator.GetITDWordTable(ws);
                    wordTables.Add(wordTable);
                }
                else if (ws.WorkSheetType == WorkSheetTypeEnum.UnderwritingYear)
                {
                    var wordTable = WordTableGenerator.GetUWYearWordTable(ws);
                    wordTables.Add(wordTable);
                }
            }


            WordHelper.ExportDocumentWithDataTables(setting, wordTables);
        }
    }
}
