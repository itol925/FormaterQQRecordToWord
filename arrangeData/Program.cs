using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace arrangeData {
    class Program {
        static void Main(string[] args) {
            ArrangeMent am = new ArrangeMent();
            am.LoadTxtFile("data.txt");
            //am.WriteTxt("data3.txt");
            //am.createWord();
            //am.createWord();
            am.WriteToWord("D://data.doc");
            Console.Read();
        }
    }

    enum LineType { 
        CONTENT = 0,
        DATE = 1,
        TIME = 2,
        QQNAME = 3
    }
    enum RecordType {
        QQRecord = 0,
        DateRecord = 1,
        OtherRecord = 2
    }
    class Record { 
        public string QQName = "";
        public string time = "";
        public string content = "";
        public RecordType recordType = RecordType.QQRecord;
    }
    class ArrangeMent { 
        List<string> lines = new List<string>();
        public List<string> LoadTxtFile(string fileName) {
            lines.Clear();
            StreamReader sr = new StreamReader(fileName, Encoding.Default);
            while(sr.Peek() >= 0){
                string line = sr.ReadLine();
                if(line.Length > 0){
                    lines.Add(line);
                }
            }
            sr.Close();
            Console.WriteLine("LoadTxtFile. line count:" + lines.Count);
            return lines;
        }

        public void WriteTxt(string fileName) {
            Record cr = null;
            List<Record> records = new List<Record>();
            for (int i = 0; i < lines.Count; i++) {
                LineType lt = GetLineType(i);
                if (lt == LineType.DATE) {
                    cr = new Record();
                    records.Add(cr);

                    cr.content = lines[i];
                    cr.recordType = RecordType.DateRecord;
                } else if (lt == LineType.QQNAME) {
                    cr = new Record();
                    records.Add(cr);

                    cr.QQName = lines[i];
                    cr.recordType = RecordType.QQRecord;
                } else if (lt == LineType.TIME) {
                    cr.time = lines[i];
                } else if (lt == LineType.CONTENT) {
                    if (cr == null) { 
                        cr = new Record();
                        cr.recordType = RecordType.OtherRecord;
                    }
                    string line = lines[i];
                    while (line.Length > 0 && (line[0] == ' ' || line[line.Length - 1] == ' ')) { 
                        line = line.Trim();
                    }
                    cr.content = line;
                }
                if (i % 100 == 0) { 
                    Console.WriteLine("lines:" + i + "/" + lines.Count);
                }
            }

            string content = "";
            string newLine = "\r\n";
            string lastQQname = "";
            for (int i = 0; i < records.Count; i++) { 
                Record r = records[i];
                if (r.recordType == RecordType.DateRecord) {
                    content += newLine + "-------------------------- " + r.content + " --------------------------" + newLine;
                } else if(r.recordType == RecordType.QQRecord){
                    if (r.content == "") { 
                        continue;
                    }
                    if(r.QQName != lastQQname)
                        content += newLine + r.QQName + "    " + r.time + newLine;
                    content += r.content + newLine;

                    lastQQname = r.QQName;
                } else { 
                    content += r.content + newLine;
                }

                if (i % 100 == 0) { 
                    Console.WriteLine("records:" + i + "/" + records.Count);
                }
            }
            StreamWriter sw = new StreamWriter(fileName);
            sw.Write(content);
            sw.Flush();
            sw.Close();
            Console.WriteLine("over");
        }

        public LineType GetLineType(int index) {
            if (index >= lines.Count - 1) { 
                return LineType.CONTENT;
            }
            string line = lines[index];
            if (IsTime(line)) { 
                return LineType.TIME;
            }
            if (IsDate(line)) { 
                return LineType.DATE;
            }
            string nextLine = lines[index + 1];
            if (IsTime(nextLine)) { 
                return LineType.QQNAME;
            }
            return LineType.CONTENT;
        }
        public bool IsTime(string line) { 
            return Regex.IsMatch(line, @"^((20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d)$");
        }
        public bool IsDate(string line)
        {
            if (line.IndexOf("日期:") != 0 || line.Length < 11) { 
                return false;
            }
            line = line.Substring(3).Trim();
            return Regex.IsMatch(line, @"^((((1[6-9]|[2-9]\d)\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\d|3[01]))|(((1[6-9]|[2-9]\d)\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\d|30))|(((1[6-9]|[2-9]\d)\d{2})-0?2-(0?[1-9]|1\d|2[0-9]))|(((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-))$");
        }

        public void WriteToWord(string filename) {
            if (File.Exists(filename)) { 
                File.Delete(filename);
            }
            Record cr = null;
            List<Record> records = new List<Record>();

            #region construct records
            for (int i = 0; i < lines.Count; i++) {
                LineType lt = GetLineType(i);
                if (lt == LineType.DATE) {
                    cr = new Record();
                    records.Add(cr);

                    cr.content = lines[i];
                    cr.recordType = RecordType.DateRecord;
                } else if (lt == LineType.QQNAME) {
                    cr = new Record();
                    records.Add(cr);

                    cr.QQName = lines[i];
                    cr.recordType = RecordType.QQRecord;
                } else if (lt == LineType.TIME) {
                    cr.time = lines[i];
                } else if (lt == LineType.CONTENT) {
                    if (cr == null) { 
                        cr = new Record();
                        cr.recordType = RecordType.OtherRecord;
                    }
                    string line = lines[i];
                    while (line.Length > 0 && (line[0] == ' ' || line[line.Length - 1] == ' ')) { 
                        line = line.Trim();
                    }
                    cr.content = line;
                }
                //if(i > 1000)break;
                if (i % 100 == 0) { 
                    Console.WriteLine("lines:" + i + "/" + lines.Count);
                }
            }
            #endregion
                        
            //创建一个document.
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";     //endofdoc是预定义的bookmark
            Microsoft.Office.Interop.Word._Application oWord;
            Microsoft.Office.Interop.Word._Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            //oWord.Visible = false;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);


            string lastQQname = "";
            for (int i = 0; i < records.Count; i++) { 
                Record r = records[i];
                if (r.recordType == RecordType.DateRecord) {
                    string content = "------------- " + r.content + " -------------";
                    addWordParagraph(oDoc, oMissing, content, 1, WdColor.wdColorBlack, 20, 20, 16);
                } else if(r.recordType == RecordType.QQRecord){
                    if (r.content == "") { 
                        continue;
                    }
                    if (r.QQName != lastQQname) {
                        string content = r.QQName + "    " + r.time;
                        addWordParagraph(oDoc, oMissing, content, 1, WdColor.wdColorBlue, 1, 12, 11, "隶书");
                    }
                    
                    string content1 = r.content;
                    addWordParagraph(oDoc, oMissing, content1, 1, WdColor.wdColorGray625, 1, 1, 10);

                    lastQQname = r.QQName;
                } else { 
                    string content = r.content;
                    addWordParagraph(oDoc, oMissing, content, 1, WdColor.wdColorBlack, 5, 5, 10);
                }

                if (i % 10 == 0) { 
                    Console.WriteLine("records:" + i + "/" + records.Count);
                }
            }
            object fn = filename;
            //oDoc.Save();
            oDoc.SaveAs(ref fn, ref oMissing, ref oMissing, ref oMissing, ref oMissing, 
                ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing,
                ref oMissing,ref oMissing,ref oMissing,ref oMissing,ref oMissing);
 
            oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            Console.WriteLine("over");
        }
        public void addWordParagraph(_Document oDoc, object oMissing, 
            string text, int bold, WdColor color, float spaceAfter, float spaceBefore, float fsize, string fontName = "宋体") { 

            Microsoft.Office.Interop.Word.Paragraph oPara = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara.Range.Text = text;
            oPara.Range.Font.Size = fsize;
            //oPara.Range.Font.Bold = bold;
            oPara.Range.Font.Name = fontName;
            oPara.Range.Font.Color = color;
            oPara.Format.SpaceAfter = spaceAfter;        // 行间距
            oPara.Format.SpaceBefore = spaceBefore;
            oPara.Range.InsertParagraphAfter();
        }
        public void createWord() { 
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";     //endofdoc是预定义的bookmark
 
            //创建一个document.
            Microsoft.Office.Interop.Word._Application oWord;
            Microsoft.Office.Interop.Word._Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
 
            //在document的开始部分添加一个paragraph.
            Microsoft.Office.Interop.Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Heading 1";
            oPara1.Range.Font.Bold = 1;
            oPara1.Range.Font.Color = WdColor.wdColorBlue;
            oPara1.Format.SpaceAfter = 24;        //24 pt 行间距
            oPara1.Range.InsertParagraphAfter();
 
            //在当前document的最后添加一个paragraph
            Microsoft.Office.Interop.Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Heading 2";
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();
 
            //接着添加一个paragraph
            Microsoft.Office.Interop.Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 24;
            oPara3.Range.InsertParagraphAfter();
 
            //添加一个3行5列的表格，填充数据，并且设定第一行的样式
            Microsoft.Office.Interop.Word.Table oTable;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;
            string strText;
            for (r = 1; r <= 3; r++)
                for (c = 1; c <= 5; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;
 
            //接着添加一些文字
            Microsoft.Office.Interop.Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "And here's another table:";
            oPara4.Format.SpaceAfter = 24;
            oPara4.Range.InsertParagraphAfter();
 
            //添加一个5行2列的表，填充数据并且改变列宽
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            for (r = 1; r <= 5; r++)
                for (c = 1; c <= 2; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Columns[1].Width = oWord.InchesToPoints(2); //设置列宽
            oTable.Columns[2].Width = oWord.InchesToPoints(3);
 
            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            object oPos;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 6;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                                           (Microsoft.Office.Interop.Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));
            object oCollapseEnd = Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();
 
            //添加一个chart
            Microsoft.Office.Interop.Word.InlineShape oShape;
            object oClassType = "MSGraph.Chart.8";
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);
 
            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            object oChart;
            object oChartApp;
            oChart = oShape.OLEFormat.Object;
            oChartApp = oChart.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oChart, null);
 
            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
            null, oChart, Parameters);
 
            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update",
            BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit",
            BindingFlags.InvokeMethod, null, oChartApp, null);
            //... If desired, you can proceed from here using the Microsoft Graph
            //Object model on the oChart and oChartApp objects to make additional
            //changes to the chart.
 
            //Set the width of the chart.
            oShape.Width = oWord.InchesToPoints(6.25f);
            oShape.Height = oWord.InchesToPoints(3.57f);
 
            //Add text after the chart.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("THE END.");
 
            Console.ReadLine();
        }
    }
}
