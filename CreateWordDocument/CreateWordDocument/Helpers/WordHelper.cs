using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class WordCell
    {
        public WordCell()
        {
            FontSize = 10;
            FontFamily = "Calibri";
            FontBold = false;
            H_Alignment = ParagraphAlignment.LEFT;
            V_Alignment = TextAlignment.CENTER;
            Content = string.Empty;
        }

        private int mergeColumnNumber;
        public int MergeColumnNumber
        {
            get
            {
                if (mergeColumnNumber <= 0)
                    return 1;
                else
                    return mergeColumnNumber;
            }
            set
            {
                mergeColumnNumber = value;
            }
        }
        public int Width { get; set; }
        public string Content { get; set; }
        public int FontSize { get; set; }
        public string FontFamily { get; set; }
        public bool FontBold { get; set; }
        public ParagraphAlignment H_Alignment { get; set; }
        public TextAlignment V_Alignment { get; set; }
    }
    public class WordTable
    {
        public WordTable(int columnCount)
        {
            this.ColumnCount = columnCount;
            //this.Width = width;
        }

        public WordTable(int columnCount, List<List<WordCell>> rows)
        {
            this.ColumnCount = columnCount;
            //this.Width = width;
            //this.ColumnWidth = columnWidth;
            this.Rows = rows;
        }
        public int ColumnCount { get; set; }
        public int Width { get; set; }
        public List<int> ColumnWidth { get; set; }
        public List<List<WordCell>> Rows { get; set; }

        public List<WordCell> CurrentRow
        {
            get
            {
                if (Rows == null)
                    throw new Exception("row index out of range");

                var currentRow = Rows.LastOrDefault();

                return currentRow;
            }
        }

        public List<WordCell> AddRow()
        {
            if (this.Rows == null)
            {
                Rows = new List<List<WordCell>>();
            }

            var newRow = new List<WordCell>();
            for (var i = 0; i < ColumnCount; i++)
            {
                var newCell = new WordCell();
                newRow.Add(newCell);
            }

            Rows.Add(newRow);

            return newRow;
        }

        public WordCell GetCell(int row, int col)
        {
            if (Rows == null || Rows.Count < row + 1)
                return null;

            var currentRow = Rows[row];

            if (currentRow == null || currentRow.Count < col + 1)
                return null;

            return currentRow[col];
        }

        public List<WordCell> GetRow(int row)
        {
            if (Rows == null || Rows.Count < row + 1)
                return null;

            var currentRow = Rows[row];

            return currentRow;
        }

        public void Merge(int row, int col, int mergeNumber)
        {
            var currentRow = GetRow(row);

            if (currentRow == null)
                throw new Exception("row index out of range");

            if(col + 1 + mergeNumber > ColumnCount)
                throw new Exception("column index out of range");

            for (var i = col + 1; i < col + 1 + mergeNumber; i++)
            {
                var needRemoveCell = GetCell(row, i);

                if (needRemoveCell != null)
                    currentRow.Remove(needRemoveCell);
            }
        }

        public void MergeCurrentRow(int col, int mergeNumber)
        {
            if (CurrentRow == null)
                throw new Exception("row index out of range");

            if (col + 1 + mergeNumber > ColumnCount)
                throw new Exception("column index out of range");

            CurrentRow[col].MergeColumnNumber = mergeNumber + 1;

            for (var i = col + 1; i < col + 1 + mergeNumber; i++)
            {
                var needRemoveCell = CurrentRow[i];

                if (needRemoveCell != null)
                    CurrentRow.Remove(needRemoveCell);
            }
        }

        public void AdjustColumnWidth()
        {
            if(ColumnWidth == null || ColumnWidth.Count != ColumnCount || Rows == null || Rows.Count == 0)
                throw new Exception("column index out of range");

            for (var i = 0; i < ColumnWidth.Count; i++)
            {
                var w = ColumnWidth[i];

                foreach(var row in Rows)
                {
                    if (row == null || row.Count == 0)
                        continue;

                    for (var c = 0; c < row.Count; c++)
                    {
                        if(c == i)
                        {
                            row[c].Width = w;
                        }
                    }
                }
            }
        }
    }

    public class DocumentSetting
    {
        public PaperType PaperType { get; set; }
        public PaperMarType PaperMarType { get; set; } 
        public string SavePath { get; set; }
        public string Footer { get; set; }
    }
    
    public enum PaperMarType
    {
        MarType1
    }
    /// <summary>
    /// 
    /// </summary>
    public enum PaperType
    {
        /// <summary>
        /// A4 Vertical
        /// </summary>
        A4_V,
        /// <summary>
        /// A4 Horizontal
        /// </summary>
        A4_H,

        /// <summary>
        /// A5 V
        /// </summary>
        A5_V,
        /// <summary>
        /// A5 H
        /// </summary>
        A5_H,

        /// <summary>
        /// A6 V
        /// </summary>
        A6_V,
        /// <summary>
        /// A6 H
        /// </summary>
        A6_H
    }

    public class WordHelper
    {
        public static XWPFParagraph SetCellText(XWPFDocument doc, XWPFTable table, WordCell wordCell)
        {
            CT_P para = new CT_P();
            XWPFParagraph pCell = new XWPFParagraph(para, table.Body);

            pCell.Alignment = wordCell.H_Alignment;  
            pCell.VerticalAlignment = wordCell.V_Alignment;  

            XWPFRun r1c1 = pCell.CreateRun();
            r1c1.SetText(wordCell.Content);
            r1c1.FontSize = wordCell.FontSize;
            r1c1.FontFamily = wordCell.FontFamily;
            r1c1.IsBold = wordCell.FontBold;
            return pCell;
        }

        public static void ExportDocumentWithDataTables(DocumentSetting setting, List<WordTable> wordTables)
        {
            XWPFDocument docx = new XWPFDocument();
            MemoryStream ms = new MemoryStream();

            #region Page Setting
            //Set Document
            docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr setPr = docx.Document.body.sectPr;
            //get page size
            Tuple<int, int> size = GetPaperSize(setting.PaperType);
            setPr.pgSz.w = (ulong)size.Item1;
            setPr.pgSz.h = (ulong)size.Item2;

            var mar = GetPaperMar(setting.PaperMarType);

            setPr.pgMar.left = (ulong)mar.Item3;
            setPr.pgMar.right = (ulong)mar.Item4;
            setPr.pgMar.top = mar.Item1.ToString();
            setPr.pgMar.bottom = mar.Item2.ToString();

            #region Footer
            CT_Ftr m_ftr = new CT_Ftr();

            m_ftr.AddNewP().AddNewR().AddNewT().Value = setting.Footer;
             
            XWPFRelation Frelation = XWPFRelation.FOOTER;
            XWPFFooter m_f = (XWPFFooter)docx.CreateRelationship(Frelation, XWPFFactory.GetInstance(), docx.FooterList.Count + 1);
           
            m_f.SetHeaderFooter(m_ftr);
            CT_HdrFtrRef m_HdrFtr1 = setPr.AddNewFooterReference();

            m_HdrFtr1.type = ST_HdrFtr.@default;
            m_HdrFtr1.id = m_f.GetPackageRelationship().Id;
            #endregion

            #endregion Page Setting

            var i = 0;
            foreach (var wordTable in wordTables)
            {
                i++;
                ExportDocumentWithDataTable(docx, setting, wordTable);

                if (i < wordTables.Count)
                {
                    XWPFParagraph m_xp = docx.CreateParagraph();

                    m_xp.CreateRun().AddBreak();
                }
            }

            docx.Write(ms);

            using (FileStream fs = new FileStream(setting.SavePath, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
            ms.Close();
        }

        public static void ExportDocumentWithDataTable(XWPFDocument docx, DocumentSetting setting, WordTable wordTable)
        {
            if (wordTable == null || wordTable.Rows == null ||  wordTable.Rows.Count == 0)
                return;

            CT_P p = docx.Document.body.AddNewP();

            p.AddNewPPr().AddNewJc().val = ST_Jc.center;

            var Rows = wordTable.Rows;

            XWPFTable table = docx.CreateTable(1, wordTable.ColumnCount);
            
            table.RemoveRow(0);//remove first blank row
            table.Width = wordTable.Width;

            for (var i = 0; i < Rows.Count; i++)
            {
                var row = Rows[i];
                if (row == null || row.Count == 0)
                    continue;

                CT_Row nr = new CT_Row();
                XWPFTableRow mr = new XWPFTableRow(nr, table); 
                table.AddRow(mr);

                for (var j = 0; j < row.Count; j++)
                {
                    var cell = row[j];

                    var c1 = mr.CreateCell();
                    var ct = c1.GetCTTc();
                    var cp = ct.AddNewTcPr();
                    cp.gridSpan = new CT_DecimalNumber();
                    cp.gridSpan.val = Convert.ToString(cell.MergeColumnNumber);

                    var tblW = cp.AddNewTcW();
                    tblW.type = ST_TblWidth.dxa;
                    tblW.w = cell.Width.ToString();

                    c1.SetParagraph(SetCellText(docx, table, cell));
                   
                    c1.SetBorderTop(XWPFTable.XWPFBorderType.NONE,0,0,"#FFFFFF");
                    c1.SetBorderRight(XWPFTable.XWPFBorderType.NONE,0,0,"#FFFFFF");
                    c1.SetBorderLeft(XWPFTable.XWPFBorderType.NONE,0,0,"#FFFFFF");
                    c1.SetBorderBottom(XWPFTable.XWPFBorderType.NONE,0,0,"#FFFFFF");
                    
                }
            }
           
        }

        #region private
        /// <summary>
        /// up down left right
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static Tuple<int, int, int, int> GetPaperMar(PaperMarType type)
        {
            Tuple<int, int, int, int> res = null;
            switch (type)
            {
                case PaperMarType.MarType1:
                    res = new Tuple<int, int, int, int>((int)(6.3*56.7), (int)(11.1 * 56.7), (int)(14.3 * 56.7), (int)(25.4 * 56.7));
                    break;
            }

            return res;
        }
        private static Tuple<int, int> GetPaperSize(PaperType type)
        {
            Tuple<int, int> res = null;
            switch (type)
            {
                case PaperType.A4_V:
                    res = new Tuple<int, int>(11906, 16838);
                    break;
                case PaperType.A4_H:
                    res = new Tuple<int, int>(16838, 11906);
                    break;

                case PaperType.A5_V:
                    res = new Tuple<int, int>(8390, 11906);
                    break;
                case PaperType.A5_H:
                    res = new Tuple<int, int>(11906, 8390);
                    break;

                case PaperType.A6_V:
                    res = new Tuple<int, int>(5953, 8390);
                    break;
                case PaperType.A6_H:
                    res = new Tuple<int, int>(8390, 5953);
                    break;
            }
            return res;
        }
        #endregion
    }
}
