using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class ExcelGenerator
    {
        public List<string> HeaderNameList { get; set; }
        public List<string> PropertyNameList { get; set; }
        public List<ExcelDataDto> ExcelDataList { get; set; }

        public string FileName { get; set; }
        public string FileFullName
        {
            get
            {
                return GlobalConstants.GenerateFolderPath + FileName;
            }
        }

        public ExcelGenerator()
        {

        }

        public void Generate()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");

            var helper = new ExcelHelper();

            helper.CreateSheetByObjectList<ExcelDataDto>(workbook, sheet, ExcelDataList, HeaderNameList, PropertyNameList);

            FileStream streamFile = new FileStream(FileFullName, FileMode.Create);

            System.IO.MemoryStream streamMemory = new System.IO.MemoryStream();
            workbook.Write(streamMemory);
            byte[] data = streamMemory.ToArray();

            streamFile.Write(data, 0, data.Length);

            streamFile.Flush();
            streamFile.Close();
            workbook = null;
            streamMemory.Close();
            streamMemory.Dispose();
        }
    }
}
