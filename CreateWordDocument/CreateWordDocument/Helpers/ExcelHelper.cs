using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class ExcelHelper
    {
        public Dictionary<string,string> CreateFormatDictionary(string properties, string formats)
        {
            var ps = properties.Split(',').ToList();
            var fs = formats.Split(',').ToList();
            var o = new Dictionary<string, string>();
          
            var i = 0;
            foreach (var p in ps)
            {
                var f = fs[i];
                o.Add(p, f);
                i++;
            }
          
            return o;
        }

        public void CreateSheetByObjectList<T>(IWorkbook myworkbook,ISheet sheet, List<T> list, List<string> headerNameList, List<string> propertyNameList, Dictionary<string, string> formatDictionary = null)
        {
            ICellStyle styleHeader = myworkbook.CreateCellStyle();
            styleHeader.Alignment = HorizontalAlignment.Center;

            styleHeader.FillPattern = FillPattern.SolidForeground;
            styleHeader.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            IFont fontHeader = myworkbook.CreateFont(); 
            //fontHeader.FontName = "Calibri"; 
            fontHeader.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
            fontHeader.FontHeightInPoints = 11;
           
            styleHeader.SetFont(fontHeader); 

            var properties = typeof(T).GetProperties();
            var startRowIndex = 0;
            var startColIndex = 0;

            var headerRow = sheet.CreateRow(startRowIndex);

            int headeri = startColIndex;
            foreach (var n in headerNameList)
            {
                var cell = headerRow.CreateCell(headeri);
                cell.SetCellValue(n);

                cell.CellStyle = styleHeader;
                headeri++;
            }
            
            startRowIndex++;

            if (list == null)
                return;

            if (formatDictionary == null)
                formatDictionary = CreateFormatDictionary("ExchangeRate","N6");

            var extraCellStyles = new Dictionary<string, ICellStyle>();
            foreach(var dic in formatDictionary)
            {
                if(!extraCellStyles.ContainsKey(dic.Value) && !string.IsNullOrEmpty(dic.Value))
                {
                    ICellStyle cellStyle1 = myworkbook.CreateCellStyle();
                    cellStyle1.Alignment = HorizontalAlignment.Right;
                    var formatStr = GetCellFormat(dic.Value);
                    var format = myworkbook.CreateDataFormat();
                    cellStyle1.DataFormat = format.GetFormat(formatStr);
                    extraCellStyles.Add(dic.Value, cellStyle1);
                }
            }

            var cellStyles = GetFieldTypeCellStyles(myworkbook);

            foreach (var o in list)
            {
                var row = sheet.CreateRow(startRowIndex);
                var i = startColIndex;

                foreach (var pn in propertyNameList)
                {
                    foreach (var p in properties)
                    {
                        if (p.Name == pn)
                        {
                            var format = string.Empty;

                            if (formatDictionary.ContainsKey(p.Name))
                                format = formatDictionary[p.Name];

                            ICellStyle cellStyle;

                            if (!string.IsNullOrEmpty(format))
                            {
                                cellStyle = extraCellStyles[format];
                            }
                            else
                            {
                                var fieldType = GetFieldTypeByPropertyType(p.PropertyType);
                                if (!cellStyles.TryGetValue(fieldType, out cellStyle))
                                    cellStyle = cellStyles[FieldType.String];
                            }
                            var cell = row.CreateCell(i);
                            SetCell(cell, p, o);

                            cell.CellStyle = cellStyle;
                            break;
                        }
                        else
                        {
                            var cell = row.CreateCell(i);
                        }
                    }

                    i++;
                }
              
                startRowIndex++;
            }

            AutoColumnWidth(sheet, propertyNameList.Count());
        }

        public void AutoColumnWidth(ISheet sheet, int cols)
        {
            for (int col = 0; col < cols; col++)
            {
                sheet.AutoSizeColumn(col);
                int columnWidth = sheet.GetColumnWidth(col) / 256;
                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    ICell cell = row.GetCell(col);
                    int contextLength = Encoding.UTF8.GetBytes(cell.ToString()).Length;
                    columnWidth = columnWidth < contextLength ? contextLength : columnWidth;

                }
                sheet.SetColumnWidth(col, columnWidth * 200);

            }
        }

        private string GetCellFormat(string format, int len = 2)
        {
            var formatOriginal = format;

            var result = string.Empty;

            if(formatOriginal.Length==2 && (formatOriginal.StartsWith("N") || formatOriginal.StartsWith("C")))
            {
                var lenStr = formatOriginal.Substring(1, 1);

                if(int.TryParse(lenStr, out len))
                {
                    format = formatOriginal.Substring(0, 1);
                }
            }

            var decimalStr = "";

            for(int i = 0; i < len; i++)
            {
                decimalStr += "0";
            }

            if(!string.IsNullOrEmpty(decimalStr))
            {
                decimalStr = "." + decimalStr;
            }

            switch (format)
            {
                case "C":
                    result = "$#,##0" + decimalStr;
                    break;
                case "N":
                    result = "#,##0" + decimalStr;
                    break;
                default:
                    result = format;
                    break;
            }
            return result;
        }
        private FieldType GetFieldTypeByPropertyType(Type type)
        {
            var fieldType = new FieldType();

            if (type == typeof(DateTime) || type == typeof(DateTime?))
            {
                fieldType = FieldType.Date;
            }

            if (type == typeof(int) || type == typeof(int?))
            {
                fieldType = FieldType.Int;
            }
            if(type == typeof(long) || type == typeof(long?))
            {
                fieldType = FieldType.Long;
            }
            if (type == typeof(decimal) || type == typeof(decimal?)
                || type == typeof(double) || type == typeof(double?))
            {
                fieldType = FieldType.Decimal2;
            }
            if (type == typeof(string))
            {
                fieldType = FieldType.String;
            }

            return fieldType;
        }
        private void SetCell<T>(ICell cell, PropertyInfo p, T o)
        {
            var type = p.PropertyType;

            if (type == typeof(DateTime) || type == typeof(DateTime?))
            {
                var d = ((DateTime?)p.GetValue(o));
                
                if (d != null)
                    cell.SetCellValue(d.Value);
            }

            if (type == typeof(int) || type == typeof(int?))
            {
                var d = ((int?)p.GetValue(o));

                if (d != null)
                    cell.SetCellValue(d.Value);
            }

            if (type == typeof(long) || type == typeof(long?))
            {
                var d = ((long?)p.GetValue(o));

                if (d != null)
                    cell.SetCellValue(d.Value);
            }
            
            if (type == typeof(decimal) || type == typeof(decimal?) 
                || type == typeof(double) || type == typeof(double?))
            {
                double? d = null;

                if (type == typeof(decimal) || type == typeof(decimal?))
                {
                    var dd = ((decimal?)p.GetValue(o));
                    d = (double?)dd;
                }
                else
                    d = ((double?)p.GetValue(o));
                
                if (d != null)
                {
                    cell.SetCellValue(d.Value);
                }
            }
            if (type == typeof(string))
            {
                var d = ((string)p.GetValue(o));
                cell.SetCellValue(d);
            }
        }

        #region Create sheet without reflector
        public void CreateSheetByValuesList(IWorkbook myworkbook, ISheet sheet, List<List<Tuple<string, string, FieldType>>> list)
        {
            if (list == null || list.Count == 0)
                return;

            ICellStyle styleHeader = myworkbook.CreateCellStyle();
            styleHeader.Alignment = HorizontalAlignment.Center;

            styleHeader.FillPattern = FillPattern.SolidForeground;
            styleHeader.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            IFont fontHeader = myworkbook.CreateFont();
            //fontHeader.FontName = "Calibri"; 
            fontHeader.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
            fontHeader.FontHeightInPoints = 11;

            styleHeader.SetFont(fontHeader);

            var startRowIndex = 0;
            var startColIndex = 0;

            var headerRow = sheet.CreateRow(startRowIndex);

            int headeri = startColIndex;
            foreach (var n in list[0])
            {
                var cell = headerRow.CreateCell(headeri);
                cell.SetCellValue(n.Item1);

                cell.CellStyle = styleHeader;
                headeri++;
            }

            startRowIndex++;

            if (list == null)
                return;

            var cellStyles = GetFieldTypeCellStyles(myworkbook);
            foreach (var r in list)
            {
                var row = sheet.CreateRow(startRowIndex);
                var i = startColIndex;

                foreach (var col in r)
                {
                    var cell = row.CreateCell(i);
                    ICellStyle cellStyle;
                    if (!cellStyles.TryGetValue(col.Item3, out cellStyle))
                        cellStyle = cellStyles[FieldType.String];

                    SetCell(cell, col.Item3, col.Item2);
                    cell.CellStyle = cellStyle;

                    i++;
                }

                startRowIndex++;
            }

            AutoColumnWidth(sheet, list[0].Count());
        }
        private Dictionary<FieldType, ICellStyle> GetFieldTypeCellStyles(IWorkbook myworkbook)
        {
            ICellStyle cellStyle1 = myworkbook.CreateCellStyle();
            ICellStyle cellStyle2 = myworkbook.CreateCellStyle();
            ICellStyle cellStyle3 = myworkbook.CreateCellStyle();
            ICellStyle cellStyle4 = myworkbook.CreateCellStyle();
            ICellStyle cellStyle5 = myworkbook.CreateCellStyle();
            ICellStyle cellStyle6 = myworkbook.CreateCellStyle();
            ICellStyle cellStyle7 = myworkbook.CreateCellStyle();

            var formatStr = "";
            //string
            cellStyle1.Alignment = HorizontalAlignment.Left;

            //int
            cellStyle2.Alignment = HorizontalAlignment.Right;

            //decimal0
            cellStyle3.Alignment = HorizontalAlignment.Right;
            var format3 = myworkbook.CreateDataFormat();
            formatStr = GetCellFormat("N0");
            cellStyle3.DataFormat = format3.GetFormat(formatStr);

            //decimal2
            cellStyle4.Alignment = HorizontalAlignment.Right;
            var format4 = myworkbook.CreateDataFormat();
            formatStr = GetCellFormat("N2");
            cellStyle4.DataFormat = format4.GetFormat(formatStr);

            //decimal6
            cellStyle5.Alignment = HorizontalAlignment.Right;
            var format5 = myworkbook.CreateDataFormat();
            formatStr = GetCellFormat("N6");
            cellStyle5.DataFormat = format5.GetFormat(formatStr);

            //amount2
            cellStyle6.Alignment = HorizontalAlignment.Right;
            var format6 = myworkbook.CreateDataFormat();
            formatStr = GetCellFormat("C2");
            cellStyle6.DataFormat = format6.GetFormat(formatStr);

            //date
            cellStyle7.Alignment = HorizontalAlignment.Right;
            var format7 = myworkbook.CreateDataFormat();
            formatStr = "dd-MMM-yyyy";
            cellStyle7.DataFormat = format7.GetFormat(formatStr);

            var cellStyles = new Dictionary<FieldType, ICellStyle>();
            cellStyles.Add(FieldType.String, cellStyle1);
            cellStyles.Add(FieldType.Int, cellStyle2);
            cellStyles.Add(FieldType.Decimal0, cellStyle3);
            cellStyles.Add(FieldType.Decimal2, cellStyle4);
            cellStyles.Add(FieldType.Decimal6, cellStyle5);
            cellStyles.Add(FieldType.Amount2, cellStyle6);
            cellStyles.Add(FieldType.Date, cellStyle7);

            return cellStyles;
        }
        private void SetCell(ICell cell, FieldType fieldType, string v)
        {
            var formatStr = string.Empty;

            var tv1 = new DateTime();
            double tv2 = 0;
            int tv3 = 0;
            long tv4 = 0;

            switch (fieldType)
            {
                case FieldType.String:
                    cell.SetCellValue(v);
                    break;
                case FieldType.Date:
                    if(DateTime.TryParse(v, out tv1))
                        cell.SetCellValue(tv1);
                    break;
                case FieldType.Int:
                    if (int.TryParse(v, out tv3))
                        cell.SetCellValue(tv3);

                    break;
                case FieldType.Long:
                    if (long.TryParse(v, out tv4))
                        cell.SetCellValue(tv4);

                    break;
                case FieldType.Decimal0:
                    if (double.TryParse(v, out tv2))
                        cell.SetCellValue(tv2);

                    break;
                case FieldType.Decimal2:
                    if (double.TryParse(v, out tv2))
                        cell.SetCellValue(tv2);

                    break;
                case FieldType.Decimal6:
                    if (double.TryParse(v, out tv2))
                        cell.SetCellValue(tv2);

                    break;
                case FieldType.Amount2:
                    if (double.TryParse(v, out tv2))
                        cell.SetCellValue(tv2);

                    break;
            }
        }
        #endregion
    }
    #region field type enum just for export without reflection
    public enum FieldType
    {
        String,
        Date,
        Int,
        Long,
        Decimal0,
        Decimal2,
        Decimal6,
        Amount2
    }
    #endregion
}
