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
            var cmdDesc = "Please input option number to exceute:\r\n";
            cmdDesc += "1.Word\r\n";
            cmdDesc += "2.Excel\r\n";

            Console.WriteLine(cmdDesc);
            var readOption = Console.ReadLine();

            switch (readOption)
            {
                case "1":
                    var dataSets = TestData.CreateDocumentDataSets();

                    var generator = new DocumentGenerator();
                    generator.DocumentDataSets = dataSets;
                    generator.DocumentName = "DocumentSample.docx";
                    generator.Generate();
                    break;
                case "2":
                    var excelDatas = TestData.CreateExcelData();

                    var excelGenerator = new ExcelGenerator();
                    excelGenerator.ExcelDataList = excelDatas;
                    excelGenerator.FileName = "ExcelSample.xlsx";
                    excelGenerator.HeaderNameList = new List<string>() { "Name","Address","DateOfBirth", "SubjectName","TeacherName","Score" };
                    excelGenerator.PropertyNameList = new List<string>() { "Name", "Address", "DateOfBirth", "SubjectName", "TeacherName", "Score" };
                    excelGenerator.Generate();
                    break;
            }

            Console.ReadKey();
        }
    }
}
